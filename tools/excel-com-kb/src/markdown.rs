use crate::model::{Examples, MAX_SUMMARY_CHARS, Parameter};

#[derive(Debug, Clone, Default)]
pub struct Page {
    pub title: Option<String>,
    pub api_names: Vec<String>,
    pub description: Option<String>,
    pub body: String,
}

pub fn parse_page(input: &str) -> Result<Page, String> {
    let normalized = input.replace("\r\n", "\n");
    let Some(rest) = normalized.strip_prefix("---\n") else {
        return Err("missing YAML front matter".to_owned());
    };
    let Some((front_matter, body)) = rest.split_once("\n---\n") else {
        return Err("unterminated YAML front matter".to_owned());
    };
    let mut page = Page {
        body: body.to_owned(),
        ..Page::default()
    };
    let mut in_api_names = false;
    for line in front_matter.lines() {
        if let Some(value) = line.strip_prefix("title:") {
            page.title = Some(clean_scalar(value));
            in_api_names = false;
        } else if let Some(value) = line.strip_prefix("description:") {
            page.description = Some(clean_scalar(value));
            in_api_names = false;
        } else if line.trim() == "api_name:" {
            in_api_names = true;
        } else if in_api_names && line.trim_start().starts_with("- ") {
            page.api_names
                .push(clean_scalar(line.trim_start().trim_start_matches("- ")));
        } else if !line.starts_with(' ') && !line.starts_with('-') {
            in_api_names = false;
        }
    }
    if page.title.is_none() {
        page.title = body
            .lines()
            .find_map(|line| line.strip_prefix("# ").map(str::to_owned));
    }
    Ok(page)
}

pub fn classify_title(title: &str) -> Option<&'static str> {
    let lower = title.to_ascii_lowercase();
    if lower.ends_with(" object (excel)") || lower.ends_with(" collection (excel)") {
        Some("object")
    } else if lower.ends_with(" property (excel)") {
        Some("property")
    } else if lower.ends_with(" method (excel)") {
        Some("method")
    } else if lower.ends_with(" event (excel)") {
        Some("event")
    } else if lower.ends_with(" enumeration (excel)") {
        Some("enum")
    } else {
        None
    }
}

pub fn title_name(title: &str) -> String {
    title
        .split_whitespace()
        .next()
        .unwrap_or_default()
        .trim_matches('`')
        .to_owned()
}

pub fn concise_summary(page: &Page, fallback: &str) -> (String, String) {
    for candidate in page
        .description
        .iter()
        .chain(first_prose_line(&page.body).as_ref())
    {
        let clean = clean_markdown(candidate);
        if !clean.is_empty() && clean != page.title.clone().unwrap_or_default() {
            return (truncate(&clean), "source-short-description".to_owned());
        }
    }
    (
        truncate(&format!("Documentation page for {fallback}.")),
        "project-authored".to_owned(),
    )
}

pub fn property_access(body: &str) -> Option<String> {
    let first = first_prose_line(body)?;
    let lower = first.to_ascii_lowercase();
    if lower.contains("read/write") {
        Some("read-write".to_owned())
    } else if lower.contains("read-only") {
        Some("read-only".to_owned())
    } else if lower.contains("write-only") {
        Some("write-only".to_owned())
    } else {
        None
    }
}

pub fn parameters(body: &str) -> (Vec<Parameter>, Vec<String>) {
    let Some(section) = heading_section(body, "Parameters") else {
        return (Vec::new(), Vec::new());
    };
    let mut parameters = Vec::new();
    let mut warnings = Vec::new();
    let mut lines = section
        .lines()
        .filter(|line| line.trim_start().starts_with('|'));
    let Some(header) = lines.next() else {
        warnings.push("parameter section has no Markdown table".to_owned());
        return (parameters, warnings);
    };
    let header_columns: Vec<String> = table_columns(header)
        .into_iter()
        .map(|cell| cell.to_ascii_lowercase())
        .collect();
    let Some(required_index) = header_columns
        .iter()
        .position(|cell| cell.contains("required"))
    else {
        warnings.push("parameter table has no Required/Optional column".to_owned());
        return (parameters, warnings);
    };
    let type_index = header_columns
        .iter()
        .position(|cell| cell.contains("data type") || cell == "type");
    let description_index = header_columns
        .iter()
        .position(|cell| cell.contains("description"));
    for line in lines {
        let columns = table_columns(line);
        if columns.iter().all(|cell| {
            cell.chars()
                .all(|character| character == ':' || character == '-' || character.is_whitespace())
        }) {
            continue;
        }
        if columns.len() <= required_index || columns.is_empty() {
            continue;
        }
        let name = clean_markdown(&columns[0]).trim_matches('_').to_owned();
        if name.is_empty() {
            continue;
        }
        let state = columns[required_index].to_ascii_lowercase();
        let optionality = if state.contains("optional") {
            "optional"
        } else if state.contains("required") {
            "required"
        } else {
            "unknown"
        };
        let documented_type = type_index
            .and_then(|index| columns.get(index))
            .map(|value| clean_markdown(value))
            .filter(|value| !value.is_empty());
        let summary = description_index
            .and_then(|index| columns.get(index))
            .map(|value| truncate(&clean_markdown(value)))
            .filter(|value| !value.is_empty());
        parameters.push(Parameter {
            name,
            optionality: optionality.to_owned(),
            documented_type,
            summary,
        });
    }
    (parameters, warnings)
}

pub fn return_type(body: &str) -> Option<String> {
    let line = heading_section(body, "Return value")
        .and_then(first_prose_line)
        .or_else(|| {
            first_prose_line(body).filter(|line| {
                let lower = line.to_ascii_lowercase();
                lower.contains("returns")
                    || lower.contains("read-only")
                    || lower.contains("read/write")
                    || lower.contains("write-only")
            })
        })?;
    extract_link_text(&line)
        .or_else(|| extract_bold_text(&line))
        .or_else(|| {
            let clean = clean_markdown(&line);
            clean
                .strip_prefix("Returns ")
                .or_else(|| clean.strip_prefix("returns "))
                .map(|value| {
                    value
                        .split_whitespace()
                        .next()
                        .unwrap_or_default()
                        .trim_matches('.')
                        .to_owned()
                })
                .filter(|value| !value.is_empty() && value != "a" && value != "the")
        })
}

pub fn enum_values(body: &str) -> Vec<(String, Option<String>, Option<String>)> {
    let mut values = Vec::new();
    let mut header_seen = false;
    for line in body
        .lines()
        .filter(|line| line.trim_start().starts_with('|'))
    {
        let columns = table_columns(line);
        if columns.len() < 2 {
            continue;
        }
        if !header_seen {
            if columns
                .first()
                .is_some_and(|value| value.eq_ignore_ascii_case("Name"))
            {
                header_seen = true;
            }
            continue;
        }
        if columns.iter().all(|cell| {
            cell.chars()
                .all(|character| character == ':' || character == '-' || character.is_whitespace())
        }) {
            continue;
        }
        let name = clean_markdown(&columns[0]);
        if name.is_empty() {
            continue;
        }
        let value = columns
            .get(1)
            .map(|value| clean_markdown(value))
            .filter(|value| !value.is_empty());
        let summary = columns
            .get(2)
            .map(|value| truncate(&clean_markdown(value)))
            .filter(|value| !value.is_empty());
        values.push((name, value, summary));
    }
    values
}

pub fn examples(body: &str) -> Examples {
    let mut anchors = Vec::new();
    for line in body.lines() {
        let Some(heading) = line.strip_prefix("## ") else {
            continue;
        };
        if heading.to_ascii_lowercase().starts_with("example") {
            anchors.push(slug(heading));
        }
    }
    Examples {
        count: anchors.len(),
        anchors,
        copied: false,
    }
}

pub fn object_collection_item(body: &str) -> Option<String> {
    let prose = clean_markdown(body);
    for marker in ["contains all the ", "collection of ", "contains the "] {
        if let Some(rest) = prose
            .to_ascii_lowercase()
            .find(marker)
            .map(|index| &prose[index + marker.len()..])
        {
            let candidate = rest
                .split_whitespace()
                .next()?
                .trim_matches(|character: char| !character.is_alphanumeric() && character != '_');
            if !candidate.is_empty() {
                return Some(candidate.to_owned());
            }
        }
    }
    None
}

pub fn clean_markdown(input: &str) -> String {
    let mut result = String::new();
    let mut remaining = input.replace("<br/>", " ").replace("<br />", " ");
    while let Some(start) = remaining.find('[') {
        result.push_str(&remaining[..start]);
        let Some(close) = remaining[start..].find(']') else {
            break;
        };
        let close = start + close;
        if let (true, Some(end)) = (
            remaining[close..].starts_with("]("),
            remaining[close + 2..].find(')'),
        ) {
            result.push_str(&remaining[start + 1..close]);
            remaining = remaining[close + 3 + end..].to_owned();
            continue;
        }
        result.push_str(&remaining[start..=close]);
        remaining = remaining[close + 1..].to_owned();
    }
    result.push_str(&remaining);
    result = result.replace("**", "").replace(['`', '_'], "");
    let mut without_tags = String::new();
    let mut inside_tag = false;
    for character in result.chars() {
        match character {
            '<' => inside_tag = true,
            '>' => inside_tag = false,
            _ if !inside_tag => without_tags.push(character),
            _ => {}
        }
    }
    without_tags
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

fn clean_scalar(value: &str) -> String {
    value.trim().trim_matches('"').trim_matches('\'').to_owned()
}

fn first_prose_line(body: &str) -> Option<String> {
    body.lines()
        .map(str::trim)
        .find(|line| {
            !line.is_empty()
                && !line.starts_with('#')
                && !line.starts_with('[')
                && !line.starts_with('|')
                && !line.starts_with("```")
                && !line.starts_with("_")
                && !line.starts_with('-')
        })
        .map(str::to_owned)
}

fn heading_section<'a>(body: &'a str, title: &str) -> Option<&'a str> {
    let marker = format!("## {title}");
    let start = body
        .lines()
        .position(|line| line.trim().eq_ignore_ascii_case(&marker))?;
    let lines: Vec<&str> = body.lines().collect();
    let end = lines
        .iter()
        .skip(start + 1)
        .position(|line| line.starts_with("## "))
        .map(|offset| start + 1 + offset)
        .unwrap_or(lines.len());
    Some(
        &body[lines[..start + 1]
            .iter()
            .map(|line| line.len() + 1)
            .sum::<usize>()
            ..lines[..end]
                .iter()
                .map(|line| line.len() + 1)
                .sum::<usize>()
                .min(body.len())],
    )
}

fn table_columns(line: &str) -> Vec<String> {
    line.trim()
        .trim_matches('|')
        .split('|')
        .map(|cell| cell.trim().to_owned())
        .collect()
}

fn extract_link_text(input: &str) -> Option<String> {
    let start = input.find("[")?;
    let rest = &input[start + 1..];
    let end = rest.find(']')?;
    Some(clean_markdown(&rest[..end]))
}

fn extract_bold_text(input: &str) -> Option<String> {
    let start = input.find("**")? + 2;
    let end = input[start..].find("**")? + start;
    Some(clean_markdown(&input[start..end]))
}

fn truncate(value: &str) -> String {
    let trimmed = value.trim();
    if trimmed.chars().count() <= MAX_SUMMARY_CHARS {
        return trimmed.to_owned();
    }
    let mut result: String = trimmed
        .chars()
        .take(MAX_SUMMARY_CHARS.saturating_sub(1))
        .collect();
    result.push('…');
    result
}

fn slug(value: &str) -> String {
    value
        .to_ascii_lowercase()
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() {
                character
            } else {
                '-'
            }
        })
        .collect::<String>()
        .split('-')
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join("-")
}
