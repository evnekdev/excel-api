use std::collections::BTreeMap;

pub(crate) const MAX_TOPICS: usize = 64;
pub(crate) const MAX_REFRESH_BATCH: usize = 32;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum ServerPhase {
    Created,
    Started,
    Active,
    Stopping,
    CallbackRevocationPending,
    Terminated,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum ModelError {
    AlreadyStarted,
    InvalidState,
    InvalidTopic,
    DuplicateTopicId,
    TopicLimit,
    UnknownTopic,
    RefreshAlreadyRunning,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct RefreshItem {
    pub topic_id: i32,
    pub value: i32,
    pub version: u64,
}

#[derive(Debug)]
struct Topic {
    #[allow(dead_code)] // proves retained Automation topic text is owned
    components: Vec<Vec<u16>>,
    value: i32,
    version: u64,
    delivered: u64,
    generation: u64,
}

#[derive(Debug)]
pub(crate) struct ServerModel {
    phase: ServerPhase,
    generation: u64,
    topics: BTreeMap<i32, Topic>,
    notification_pending: bool,
    refresh_running: bool,
    notify_failures: u64,
}

impl ServerModel {
    pub(crate) fn new(generation: u64) -> Self {
        Self {
            phase: ServerPhase::Created,
            generation,
            topics: BTreeMap::new(),
            notification_pending: false,
            refresh_running: false,
            notify_failures: 0,
        }
    }

    pub(crate) fn phase(&self) -> ServerPhase {
        self.phase
    }

    #[cfg(test)]
    pub(crate) fn topic_count(&self) -> usize {
        self.topics.len()
    }

    pub(crate) fn start(&mut self) -> Result<(), ModelError> {
        match self.phase {
            ServerPhase::Created => {
                self.phase = ServerPhase::Started;
                Ok(())
            }
            ServerPhase::Started | ServerPhase::Active => Err(ModelError::AlreadyStarted),
            ServerPhase::Stopping
            | ServerPhase::CallbackRevocationPending
            | ServerPhase::Terminated => Err(ModelError::InvalidState),
        }
    }

    pub(crate) fn rollback_start(&mut self) {
        if self.phase == ServerPhase::Started && self.topics.is_empty() {
            self.phase = ServerPhase::Created;
            self.notification_pending = false;
            self.refresh_running = false;
        }
    }

    pub(crate) fn rollback_failed_start(&mut self, callback_revoked: bool) {
        if self.phase == ServerPhase::Stopping && self.topics.is_empty() {
            self.phase = if callback_revoked {
                ServerPhase::Created
            } else {
                ServerPhase::CallbackRevocationPending
            };
            self.notification_pending = false;
            self.refresh_running = false;
        }
    }

    pub(crate) fn connect(
        &mut self,
        topic_id: i32,
        components: Vec<Vec<u16>>,
    ) -> Result<i32, ModelError> {
        if !matches!(self.phase, ServerPhase::Started | ServerPhase::Active) {
            return Err(ModelError::InvalidState);
        }
        if !valid_topic(&components) {
            return Err(ModelError::InvalidTopic);
        }
        if self.topics.contains_key(&topic_id) {
            return Err(ModelError::DuplicateTopicId);
        }
        if self.topics.len() == MAX_TOPICS {
            return Err(ModelError::TopicLimit);
        }
        self.topics.insert(
            topic_id,
            Topic {
                components,
                value: 0,
                version: 0,
                delivered: 0,
                generation: self.generation,
            },
        );
        self.phase = ServerPhase::Active;
        Ok(0)
    }

    pub(crate) fn disconnect(&mut self, topic_id: i32) -> Result<(), ModelError> {
        if !matches!(self.phase, ServerPhase::Started | ServerPhase::Active) {
            return Err(ModelError::InvalidState);
        }
        self.topics
            .remove(&topic_id)
            .ok_or(ModelError::UnknownTopic)?;
        if self.topics.is_empty() {
            self.phase = ServerPhase::Started;
            self.notification_pending = false;
        }
        Ok(())
    }

    pub(crate) fn publish_counter_tick(&mut self) -> bool {
        if self.phase != ServerPhase::Active {
            return false;
        }
        for topic in self.topics.values_mut() {
            if topic.generation == self.generation {
                topic.value = topic.value.saturating_add(1);
                topic.version = topic.version.saturating_add(1);
            }
        }
        self.claim_notification()
    }

    pub(crate) fn claim_notification(&mut self) -> bool {
        let dirty = self
            .topics
            .values()
            .any(|topic| topic.version > topic.delivered);
        if self.phase == ServerPhase::Active
            && dirty
            && !self.notification_pending
            && !self.refresh_running
        {
            self.notification_pending = true;
            true
        } else {
            false
        }
    }

    pub(crate) fn notification_failed(&mut self) {
        if matches!(self.phase, ServerPhase::Started | ServerPhase::Active) {
            self.notification_pending = false;
            self.notify_failures = self.notify_failures.saturating_add(1);
        }
    }

    pub(crate) fn begin_refresh(&mut self) -> Result<Vec<RefreshItem>, ModelError> {
        if !matches!(self.phase, ServerPhase::Started | ServerPhase::Active) {
            return Err(ModelError::InvalidState);
        }
        if self.refresh_running {
            return Err(ModelError::RefreshAlreadyRunning);
        }
        self.refresh_running = true;
        self.notification_pending = false;
        Ok(self
            .topics
            .iter()
            .filter(|(_, topic)| topic.version > topic.delivered)
            .take(MAX_REFRESH_BATCH)
            .map(|(topic_id, topic)| RefreshItem {
                topic_id: *topic_id,
                value: topic.value,
                version: topic.version,
            })
            .collect())
    }

    pub(crate) fn finish_refresh(&mut self, batch: &[RefreshItem], succeeded: bool) {
        if succeeded {
            for item in batch {
                if let Some(topic) = self.topics.get_mut(&item.topic_id) {
                    topic.delivered = topic.delivered.max(item.version.min(topic.version));
                }
            }
        }
        self.refresh_running = false;
    }

    pub(crate) fn begin_stop(&mut self) -> bool {
        match self.phase {
            ServerPhase::Created | ServerPhase::Started | ServerPhase::Active => {
                self.phase = ServerPhase::Stopping;
                self.notification_pending = false;
                true
            }
            ServerPhase::Stopping
            | ServerPhase::CallbackRevocationPending
            | ServerPhase::Terminated => false,
        }
    }

    pub(crate) fn finish_stop(&mut self, callback_revoked: bool) {
        self.topics.clear();
        self.notification_pending = false;
        self.refresh_running = false;
        self.phase = if callback_revoked {
            ServerPhase::Terminated
        } else {
            ServerPhase::CallbackRevocationPending
        };
    }

    pub(crate) fn finish_callback_revocation(&mut self) {
        if self.phase == ServerPhase::CallbackRevocationPending {
            self.phase = ServerPhase::Terminated;
        }
    }

    pub(crate) fn heartbeat(&self) -> i32 {
        i32::from(matches!(
            self.phase,
            ServerPhase::Started | ServerPhase::Active
        ))
    }

    #[cfg(test)]
    fn topic_components(&self, topic_id: i32) -> Option<&[Vec<u16>]> {
        self.topics
            .get(&topic_id)
            .map(|topic| topic.components.as_slice())
    }
}

fn valid_topic(components: &[Vec<u16>]) -> bool {
    if components.len() != 1 {
        return false;
    }
    let Ok(name) = String::from_utf16(&components[0]) else {
        return false;
    };
    name.eq_ignore_ascii_case("COUNTER")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn counter() -> Vec<Vec<u16>> {
        vec!["COUNTER".encode_utf16().collect()]
    }

    #[test]
    fn state_topic_and_heartbeat_transitions_are_explicit() {
        let mut model = ServerModel::new(7);
        assert_eq!(model.phase(), ServerPhase::Created);
        assert_eq!(model.heartbeat(), 0);
        model.start().unwrap();
        assert_eq!(model.heartbeat(), 1);
        assert_eq!(model.connect(4, counter()), Ok(0));
        assert_eq!(model.phase(), ServerPhase::Active);
        assert_eq!(model.disconnect(4), Ok(()));
        assert_eq!(model.phase(), ServerPhase::Started);
        assert!(model.begin_stop());
        assert_eq!(model.heartbeat(), 0);
        model.finish_stop(true);
        assert_eq!(model.phase(), ServerPhase::Terminated);
        assert!(!model.begin_stop());
    }

    #[test]
    fn topic_inputs_are_owned_bounded_and_deterministic() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        let mut input = counter();
        model.connect(1, input.clone()).unwrap();
        input[0][0] = b'X' as u16;
        assert_eq!(model.topic_components(1), Some(counter().as_slice()));
        assert_eq!(
            model.connect(1, counter()),
            Err(ModelError::DuplicateTopicId)
        );
        assert_eq!(
            model.connect(2, vec!["CLOCK".encode_utf16().collect()]),
            Err(ModelError::InvalidTopic)
        );
        for id in 2..=MAX_TOPICS as i32 {
            model.connect(id, counter()).unwrap();
        }
        assert_eq!(model.connect(1000, counter()), Err(ModelError::TopicLimit));
    }

    #[test]
    fn equivalent_topics_may_have_distinct_excel_ids_and_reconnect() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        model.connect(10, counter()).unwrap();
        model.connect(11, counter()).unwrap();
        assert_eq!(model.topic_count(), 2);
        model.disconnect(10).unwrap();
        model.connect(10, counter()).unwrap();
        assert_eq!(model.topic_count(), 2);
    }

    #[test]
    fn notification_and_refresh_coalesce_without_losing_newer_updates() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        model.connect(1, counter()).unwrap();
        assert!(model.publish_counter_tick());
        assert!(!model.claim_notification());
        let first = model.begin_refresh().unwrap();
        assert_eq!(first[0].value, 1);
        assert!(!model.publish_counter_tick());
        model.finish_refresh(&first, true);
        assert!(model.claim_notification());
        let second = model.begin_refresh().unwrap();
        assert_eq!(second[0].value, 2);
        model.finish_refresh(&second, true);
        assert!(!model.claim_notification());
    }

    #[test]
    fn callback_and_payload_failures_leave_dirty_work_retryable() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        model.connect(1, counter()).unwrap();
        assert!(model.publish_counter_tick());
        model.notification_failed();
        assert!(model.claim_notification());
        let batch = model.begin_refresh().unwrap();
        model.finish_refresh(&batch, false);
        assert!(model.claim_notification());
    }

    #[test]
    fn termination_retires_topics_and_suppresses_updates() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        model.connect(1, counter()).unwrap();
        model.begin_stop();
        assert!(!model.publish_counter_tick());
        model.finish_stop(true);
        assert_eq!(model.topic_count(), 0);
        assert_eq!(model.connect(2, counter()), Err(ModelError::InvalidState));
    }

    #[test]
    fn partial_start_rolls_back_only_before_topics_exist() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        model.rollback_start();
        assert_eq!(model.phase(), ServerPhase::Created);
        model.start().unwrap();
        model.connect(1, counter()).unwrap();
        model.rollback_start();
        assert_eq!(model.phase(), ServerPhase::Active);
    }

    #[test]
    fn failed_worker_start_rolls_back_after_stop_commit() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        assert!(model.begin_stop());
        model.rollback_failed_start(true);
        assert_eq!(model.phase(), ServerPhase::Created);
        assert_eq!(model.heartbeat(), 0);
    }

    #[test]
    fn unresolved_callback_revocation_blocks_termination_state() {
        let mut model = ServerModel::new(1);
        model.start().unwrap();
        assert!(model.begin_stop());
        model.finish_stop(false);
        assert_eq!(model.phase(), ServerPhase::CallbackRevocationPending);
        assert_eq!(model.heartbeat(), 0);
        assert_eq!(model.connect(1, counter()), Err(ModelError::InvalidState));
        model.finish_callback_revocation();
        assert_eq!(model.phase(), ServerPhase::Terminated);
    }
}
