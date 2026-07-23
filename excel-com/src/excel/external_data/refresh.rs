//! Workbook-wide refresh orchestration and explicit STA polling.

use std::time::Instant;

use crate::ExcelComError;
use crate::automation::invoke;
use crate::excel::{Application, Workbook};
use crate::object_model::{MemberId, member};

use super::{ConnectionDetails, RefreshCancellationReport, RefreshWaitOptions, RefreshWaitReport};

impl Application {
    /// Waits for Excel's asynchronous query engine on the owning STA thread.
    ///
    /// Excel controls which provider operations this covers. For bounded
    /// per-workbook polling, use [`Workbook::wait_for_refresh`].
    pub fn calculate_until_async_queries_done(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(
                MemberId::new("excel.application.calculateuntilasyncqueriesdone"),
                false,
            ),
            vec![],
            false,
        )?;
        Ok(())
    }
}

impl Workbook {
    /// Requests Excel to refresh all refreshable workbook data objects.
    pub fn refresh_all(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.workbook.refreshall"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns whether any observable worksheet QueryTable is refreshing.
    ///
    /// Excel offers no universal `Workbook.Refreshing` flag; connection and
    /// provider states not exposed through a QueryTable are intentionally not claimed here.
    pub fn is_refreshing(&self) -> Result<bool, ExcelComError> {
        for sheet in self.worksheets()?.iter()? {
            for query in sheet?.query_tables()?.iter()? {
                if query?.refreshing()? {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }
    /// Best-effort cancellation over QueryTables and typed provider connections.
    pub fn cancel_all_refreshes(&self) -> Result<RefreshCancellationReport, ExcelComError> {
        let mut report = RefreshCancellationReport::default();
        for sheet in self.worksheets()?.iter()? {
            for query in sheet?.query_tables()?.iter()? {
                let query = query?;
                if query.refreshing()? {
                    query.cancel_refresh()?;
                    report.query_tables_cancelled += 1;
                }
            }
        }
        for connection in self.connections()?.iter()? {
            match connection?.details()? {
                ConnectionDetails::OleDb(value) => {
                    if value.refreshing()? {
                        value.cancel_refresh()?;
                        report.connections_cancelled += 1;
                    }
                }
                ConnectionDetails::Odbc(value) => {
                    if value.refreshing()? {
                        value.cancel_refresh()?;
                        report.connections_cancelled += 1;
                    }
                }
                _ => report.unsupported_refreshes += 1,
            }
        }
        Ok(report)
    }
    /// Polls observable QueryTable state with caller-supplied explicit bounds.
    ///
    /// No background thread is created. Sleep happens only between COM calls,
    /// after all temporary argument buffers have been released.
    pub fn wait_for_refresh(
        &self,
        options: RefreshWaitOptions,
    ) -> Result<RefreshWaitReport, ExcelComError> {
        options.validate()?;
        let start = Instant::now();
        loop {
            let remaining = remaining_queries(self)?;
            let elapsed = start.elapsed();
            if remaining == 0 {
                return Ok(RefreshWaitReport {
                    completed: true,
                    elapsed,
                    remaining_queries: 0,
                });
            }
            if elapsed >= options.timeout {
                return Ok(RefreshWaitReport {
                    completed: false,
                    elapsed,
                    remaining_queries: remaining,
                });
            }
            std::thread::sleep(
                options
                    .poll_interval
                    .min(options.timeout.saturating_sub(elapsed)),
            );
        }
    }
}

fn remaining_queries(workbook: &Workbook) -> Result<usize, ExcelComError> {
    let mut remaining = 0;
    for sheet in workbook.worksheets()?.iter()? {
        for query in sheet?.query_tables()?.iter()? {
            if query?.refreshing()? {
                remaining += 1;
            }
        }
    }
    Ok(remaining)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::Duration;
    #[test]
    fn wait_requires_explicit_nonzero_bounds() {
        assert!(
            RefreshWaitOptions {
                timeout: Duration::ZERO,
                poll_interval: Duration::from_millis(1)
            }
            .validate()
            .is_err()
        );
    }
}
