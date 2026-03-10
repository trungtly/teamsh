use ratatui::style::{Color, Modifier, Style};
use ratatui::text::Span;

/// A segment of rendered text with formatting info
#[derive(Debug, Clone)]
pub struct RichSegment {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub mention: bool,
    pub link_url: Option<String>,
    pub quote: bool,
}

impl RichSegment {
    fn plain(text: String) -> Self {
        Self { text, bold: false, italic: false, mention: false, link_url: None, quote: false }
    }
}

/// Strip HTML tags and decode common entities to plain text.
/// Handles Teams quote replies (blockquote) by prefixing with "> ".
pub fn strip_html(html: &str) -> String {
    let (quote, reply) = extract_quote(html);

    let stripped = strip_tags_plain(&if let Some(reply) = &reply {
        reply.clone()
    } else {
        html.to_string()
    });

    if let Some(q) = quote {
        let q_text = strip_tags_plain(&q);
        if !q_text.is_empty() {
            return format!("> {}\n{}", q_text, stripped);
        }
    }
    stripped
}

/// Rich strip: returns segments with formatting info for TUI rendering.
/// Returns Vec of (line_segments) where each line is a Vec<RichSegment>.
pub fn strip_html_rich(html: &str) -> Vec<Vec<RichSegment>> {
    let (quote, reply) = extract_quote(html);

    let mut all_lines = Vec::new();

    if let Some(q) = quote {
        let q_text = strip_tags_plain(&q);
        if !q_text.is_empty() {
            all_lines.push(vec![RichSegment {
                text: format!("> {}", q_text),
                bold: false, italic: false, mention: false, link_url: None, quote: true,
            }]);
        }
    }

    let source = if let Some(reply) = &reply {
        reply.clone()
    } else {
        html.to_string()
    };

    let segments = strip_tags_rich(&source);
    // Split segments into lines
    let mut current_line: Vec<RichSegment> = Vec::new();
    for seg in segments {
        if seg.text.contains('\n') {
            let parts: Vec<&str> = seg.text.split('\n').collect();
            for (i, part) in parts.iter().enumerate() {
                if !part.is_empty() {
                    current_line.push(RichSegment { text: part.to_string(), ..seg.clone() });
                }
                if i < parts.len() - 1 {
                    all_lines.push(std::mem::take(&mut current_line));
                }
            }
        } else {
            current_line.push(seg);
        }
    }
    if !current_line.is_empty() {
        all_lines.push(current_line);
    }
    if all_lines.is_empty() {
        all_lines.push(vec![RichSegment::plain(String::new())]);
    }
    all_lines
}

/// Convert RichSegments to ratatui Spans for rendering
// Catppuccin Mocha palette - matches teamsh bat theme
const COLOR_MENTION: Color = Color::Rgb(249, 226, 175);  // #f9e2af yellow
const COLOR_QUOTE: Color = Color::Rgb(108, 112, 134);    // #6c7086 overlay0
const COLOR_LINK: Color = Color::Rgb(137, 180, 250);     // #89b4fa blue
const COLOR_LINK_TEXT: Color = Color::Rgb(116, 199, 236); // #74c7ec sapphire

pub fn rich_to_spans(segments: &[RichSegment]) -> Vec<Span<'static>> {
    segments.iter().map(|seg| {
        let mut style = Style::default();
        if seg.quote {
            style = style.fg(COLOR_QUOTE).add_modifier(Modifier::ITALIC);
        }
        if seg.bold {
            style = style.add_modifier(Modifier::BOLD);
        }
        if seg.italic {
            style = style.add_modifier(Modifier::ITALIC);
        }
        if seg.mention {
            style = style.fg(COLOR_MENTION).add_modifier(Modifier::BOLD);
        }
        if let Some(url) = &seg.link_url {
            let text = seg.text.trim();
            let display = if text == url || text.is_empty() {
                url.clone()
            } else {
                format!("[{}]({})", text, url)
            };
            return Span::styled(display, style.fg(COLOR_LINK).add_modifier(Modifier::UNDERLINED));
        }
        Span::styled(seg.text.clone(), style)
    }).collect()
}

/// Extract (quote_preview, reply_text) from Teams HTML with blockquote.
fn extract_quote(html: &str) -> (Option<String>, Option<String>) {
    let bq_start = html.find("<blockquote");
    let bq_end = html.find("</blockquote>");
    if let (Some(start), Some(end)) = (bq_start, bq_end) {
        let bq_content = &html[start..end];
        let preview = if let Some(p_start) = bq_content.find("itemprop=\"preview\"") {
            let after = &bq_content[p_start..];
            let text_start = after.find('>').map(|i| p_start + i + 1);
            let text_end = after.find("</p>").or(after.find("</")).map(|i| p_start + i);
            match (text_start, text_end) {
                (Some(s), Some(e)) if e > s => Some(bq_content[s..e].to_string()),
                _ => None,
            }
        } else {
            None
        };

        let after_bq = &html[end + "</blockquote>".len()..];
        let reply = if !after_bq.trim().is_empty() {
            Some(after_bq.to_string())
        } else {
            None
        };

        (preview, reply)
    } else {
        (None, None)
    }
}

/// Plain text strip (used for search indexing and simple display).
/// Mentions are rendered as `@Name` prefix for each mention span.
/// Teams gives each word of a name its own span (often with different itemid),
/// so we just prefix each new mention span with `@`.
fn strip_tags_plain(html: &str) -> String {
    let mut result = String::with_capacity(html.len());
    let mut in_tag = false;
    let mut in_entity = false;
    let mut entity = String::new();
    let mut tag_buf = String::new();
    let mut pending_mention = false;
    let mut last_mention_id = String::new();
    let mut last_mention_end: Option<usize> = None; // position in result where last mention ended

    for ch in html.chars() {
        match ch {
            '<' => {
                in_tag = true;
                tag_buf.clear();
            }
            '>' if in_tag => {
                in_tag = false;
                if tag_buf.contains("schema.skype.com/Mention") && !tag_buf.starts_with('/') {
                    let mid = extract_attr(&tag_buf, "itemid");
                    if mid != last_mention_id {
                        pending_mention = true;
                        last_mention_id = mid;
                    }
                }
                // Insert newline for block-level HTML elements
                let tag_lower = tag_buf.to_lowercase();
                let tag_name = tag_lower.split_whitespace().next().unwrap_or("");
                let tag_name = tag_name.trim_start_matches('/');
                if matches!(tag_name, "br" | "p" | "div" | "tr" | "hr" | "li"
                    | "h1" | "h2" | "h3" | "h4" | "h5" | "h6"
                    | "table" | "blockquote" | "pre")
                {
                    if !result.ends_with('\n') {
                        result.push('\n');
                    }
                }
            }
            '&' if !in_tag => {
                if pending_mention {
                    emit_mention_prefix(&mut result, &mut last_mention_end);
                    pending_mention = false;
                }
                in_entity = true;
                entity.clear();
            }
            ';' if in_entity => {
                in_entity = false;
                decode_entity(&entity, &mut result);
            }
            _ if in_tag => {
                tag_buf.push(ch);
                // Track when mention span closes
                if tag_buf == "/span" {
                    // Mark end of mention content at current result position
                    if last_mention_id != "" {
                        last_mention_end = Some(result.len());
                    }
                }
            }
            _ if in_entity => {
                entity.push(ch);
            }
            _ => {
                if pending_mention {
                    emit_mention_prefix(&mut result, &mut last_mention_end);
                    pending_mention = false;
                }
                result.push(ch);
            }
        }
    }

    collapse_whitespace(&result)
}

/// Emit `@` or `-` depending on whether this mention immediately follows another.
/// If text since last mention end is only whitespace, replace it with `-` (same person).
fn emit_mention_prefix(result: &mut String, last_mention_end: &mut Option<usize>) {
    if let Some(end_pos) = *last_mention_end {
        let since = &result[end_pos..];
        if !since.is_empty() && since.chars().all(|c| c.is_whitespace()) {
            // Adjacent mention spans separated by whitespace = same person
            result.truncate(end_pos);
            result.push('-');
            *last_mention_end = None;
            return;
        }
    }
    // New mention - add space if needed then @
    if !result.is_empty() && !result.ends_with(|c: char| c.is_whitespace()) {
        result.push(' ');
    }
    result.push('@');
    *last_mention_end = None;
}

/// Rich text strip - returns formatting-aware segments
fn strip_tags_rich(html: &str) -> Vec<RichSegment> {
    let mut segments: Vec<RichSegment> = Vec::new();
    let mut current_text = String::new();
    let mut in_tag = false;
    let mut in_entity = false;
    let mut entity = String::new();
    let mut tag_buf = String::new();

    // Formatting state stack
    let mut bold_depth = 0u32;
    let mut italic_depth = 0u32;
    let mut in_mention = false;
    let mut mention_span_depth = 0u32; // track nested spans within mention
    let mut last_mention_id = String::new();
    let mut pending_mention_at = false;
    let mut last_mention_ended = false; // true when last mention span just closed
    let mut link_url: Option<String> = None;

    let flush = |segments: &mut Vec<RichSegment>, text: &mut String,
                 bold: bool, italic: bool, mention: bool, link: &Option<String>| {
        if !text.is_empty() {
            segments.push(RichSegment {
                text: text.clone(),
                bold, italic, mention,
                link_url: link.clone(),
                quote: false,
            });
            text.clear();
        }
    };

    for ch in html.chars() {
        match ch {
            '<' => {
                in_tag = true;
                tag_buf.clear();
            }
            '>' if in_tag => {
                in_tag = false;
                let tag_lower = tag_buf.to_lowercase();
                let is_close = tag_buf.starts_with('/');

                // Bold: <b>, <strong>
                if !is_close && (tag_lower == "b" || tag_lower == "strong") {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    bold_depth += 1;
                } else if is_close && (tag_lower == "/b" || tag_lower == "/strong") {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    bold_depth = bold_depth.saturating_sub(1);
                }
                // Italic: <i>, <em>
                else if !is_close && (tag_lower == "i" || tag_lower == "em") {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    italic_depth += 1;
                } else if is_close && (tag_lower == "/i" || tag_lower == "/em") {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    italic_depth = italic_depth.saturating_sub(1);
                }
                // Links: <a href="...">
                else if !is_close && tag_lower.starts_with("a ") {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    link_url = Some(extract_attr(&tag_buf, "href"));
                } else if is_close && tag_lower == "/a" {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, in_mention, &link_url);
                    link_url = None;
                }
                // Mentions
                else if !is_close && tag_buf.contains("schema.skype.com/Mention") {
                    let mid = extract_attr(&tag_buf, "itemid");
                    if mid != last_mention_id {
                        // Different itemid - could be new person or next word of same person
                        if in_mention {
                            flush(&mut segments, &mut current_text,
                                  bold_depth > 0, italic_depth > 0, true, &link_url);
                            pending_mention_at = true;
                        } else {
                            // Check if gap since last mention is only whitespace (same person)
                            let is_adjacent = last_mention_ended
                                && !current_text.is_empty()
                                && current_text.chars().all(|c| c.is_whitespace());
                            if is_adjacent {
                                // Same person - replace whitespace gap with dash in previous mention segment
                                if let Some(prev) = segments.last_mut() {
                                    if prev.mention {
                                        prev.text.push('-');
                                    }
                                }
                                current_text.clear();
                                pending_mention_at = false;
                            } else {
                                flush(&mut segments, &mut current_text,
                                      bold_depth > 0, italic_depth > 0, false, &link_url);
                                pending_mention_at = true;
                            }
                        }
                        in_mention = true;
                        last_mention_id = mid;
                    } else if !in_mention {
                        // Same person, continuation span (e.g. multi-word name)
                        in_mention = true;
                    }
                    mention_span_depth += 1;
                    last_mention_ended = false;
                }
                // Track span open/close for mention boundary detection
                else if !is_close && tag_lower.starts_with("span") && in_mention {
                    mention_span_depth += 1;
                }
                else if is_close && tag_lower == "/span" && in_mention {
                    mention_span_depth = mention_span_depth.saturating_sub(1);
                    if mention_span_depth == 0 {
                        // All mention spans closed - flush mention segment
                        flush(&mut segments, &mut current_text,
                              bold_depth > 0, italic_depth > 0, true, &link_url);
                        in_mention = false;
                        last_mention_ended = true;
                    }
                }
                // Any non-span, non-formatting opening tag closes mention
                else if !is_close && in_mention
                    && !tag_lower.starts_with("span")
                    && tag_lower != "b" && tag_lower != "strong"
                    && tag_lower != "i" && tag_lower != "em"
                {
                    flush(&mut segments, &mut current_text,
                          bold_depth > 0, italic_depth > 0, true, &link_url);
                    in_mention = false;
                    mention_span_depth = 0;
                    last_mention_id.clear();
                }
                // Line breaks
                else if !is_close && (tag_lower == "br" || tag_lower == "br/" || tag_lower == "br /") {
                    current_text.push('\n');
                }
                // Paragraph/div boundaries
                else if is_close && (tag_lower == "/p" || tag_lower == "/div") {
                    if !current_text.is_empty() && !current_text.ends_with('\n') {
                        current_text.push('\n');
                    }
                }
            }
            '&' if !in_tag => {
                if pending_mention_at {
                    if !current_text.is_empty() && !current_text.ends_with(|c: char| c.is_whitespace()) {
                        current_text.push(' ');
                    }
                    current_text.push('@');
                    pending_mention_at = false;
                }
                in_entity = true;
                entity.clear();
            }
            ';' if in_entity => {
                in_entity = false;
                decode_entity(&entity, &mut current_text);
            }
            _ if in_tag => {
                tag_buf.push(ch);
            }
            _ if in_entity => {
                entity.push(ch);
            }
            _ => {
                if pending_mention_at {
                    if !current_text.is_empty() && !current_text.ends_with(|c: char| c.is_whitespace()) {
                        current_text.push(' ');
                    }
                    current_text.push('@');
                    pending_mention_at = false;
                }
                current_text.push(ch);
            }
        }
    }

    // Close any open mention
    if in_mention {
        flush(&mut segments, &mut current_text,
              bold_depth > 0, italic_depth > 0, true, &link_url);
    } else {
        flush(&mut segments, &mut current_text,
              bold_depth > 0, italic_depth > 0, false, &link_url);
    }

    // Collapse runs of whitespace within each segment but preserve
    // leading/trailing spaces (they separate adjacent segments)
    for seg in segments.iter_mut() {
        seg.text = collapse_whitespace_inner(&seg.text);
    }
    // Trim only the first and last segment
    if let Some(first) = segments.first_mut() {
        first.text = first.text.trim_start().to_string();
    }
    if let Some(last) = segments.last_mut() {
        last.text = last.text.trim_end().to_string();
    }
    // Remove empty segments
    segments.retain(|s| !s.text.is_empty());

    segments
}

/// Extract attribute value from a tag buffer, e.g. extract_attr("a href=\"https://...\"", "href")
fn extract_attr(tag: &str, attr: &str) -> String {
    let search = format!("{}=\"", attr);
    if let Some(start) = tag.find(&search) {
        let after = &tag[start + search.len()..];
        if let Some(end) = after.find('"') {
            return after[..end].to_string();
        }
    }
    // Try single quotes
    let search = format!("{}='", attr);
    if let Some(start) = tag.find(&search) {
        let after = &tag[start + search.len()..];
        if let Some(end) = after.find('\'') {
            return after[..end].to_string();
        }
    }
    String::new()
}

fn decode_entity(entity: &str, output: &mut String) {
    match entity {
        "amp" => output.push('&'),
        "lt" => output.push('<'),
        "gt" => output.push('>'),
        "quot" => output.push('"'),
        "nbsp" => output.push(' '),
        "apos" => output.push('\''),
        _ => {
            output.push('&');
            output.push_str(entity);
            output.push(';');
        }
    }
}

fn collapse_whitespace(s: &str) -> String {
    let mut collapsed = String::with_capacity(s.len());
    let mut last_was_space = false;
    for ch in s.trim().chars() {
        if ch == '\n' {
            collapsed.push('\n');
            last_was_space = false;
        } else if ch.is_whitespace() {
            if !last_was_space {
                collapsed.push(' ');
                last_was_space = true;
            }
        } else {
            collapsed.push(ch);
            last_was_space = false;
        }
    }
    collapsed.trim().to_string()
}

/// Collapse runs of whitespace but preserve leading/trailing spaces.
fn collapse_whitespace_inner(s: &str) -> String {
    let mut collapsed = String::with_capacity(s.len());
    let mut last_was_space = false;
    for ch in s.chars() {
        if ch == '\n' {
            collapsed.push('\n');
            last_was_space = false;
        } else if ch.is_whitespace() {
            if !last_was_space {
                collapsed.push(' ');
                last_was_space = true;
            }
        } else {
            collapsed.push(ch);
            last_was_space = false;
        }
    }
    collapsed
}

/// Quick tag stripping for search indexing (no quote handling)
pub fn strip_tags_only(html: &str) -> String {
    strip_tags_plain(html)
}

/// Map Teams emoji key to Unicode
pub fn teams_emoji(key: &str) -> String {
    let base = key.split(';').next().unwrap_or(key);
    let name = if let Some(pos) = base.rfind("-tone") {
        &base[..pos]
    } else {
        base
    };

    let result = match name {
        "like" | "1f44d" | "1f44d_thumbsup" | "yes" => "\u{1f44d}",
        "heart" | "2764" | "2764_heart" => "\u{2764}",
        "laugh" | "1f606" | "1f606_laugh" => "\u{1f606}",
        "surprised" | "1f62e" | "1f62e_surprised" => "\u{1f62e}",
        "sad" | "1f622" | "1f622_sad" => "\u{1f622}",
        "angry" | "1f620" | "1f620_angry" => "\u{1f620}",
        "1f389_partypopper" | "1f389" => "\u{1f389}",
        "1f4af" | "1f4af_100" => "\u{1f4af}",
        "1f525" | "1f525_fire" => "\u{1f525}",
        "1f44f" | "1f44f_clap" | "clappinghands" | "clapclap" | "clapclap-e" => "\u{1f44f}",
        "1f64f" | "1f64f_pray" => "\u{1f64f}",
        "1f680" | "1f680_rocket" | "launch" => "\u{1f680}",
        "2705" | "2705_check" | "2714_heavycheckmark" => "\u{2705}",
        "handsinair" | "1f64c" | "1f64c_handsinair" => "\u{1f64c}",
        "1f440_eyes" => "\u{1f440}",
        "muscle" => "\u{1f4aa}",
        "bow" => "\u{1f647}",
        "loudlycrying" => "\u{1f62d}",
        "heartpurple" => "\u{1f49c}",
        "fistbump" => "\u{1f91c}",
        "stareyes" | "starstruck" | "1f929" => "\u{1f929}",
        "thinking" | "1f914" => "\u{1f914}",
        "rofl" | "1f923" => "\u{1f923}",
        "eyes" | "1f440" => "\u{1f440}",
        "pray" | "foldedhands" => "\u{1f64f}",
        "tada" | "partypopper" => "\u{1f389}",
        _ => {
            let hex_part = name.split('_').next().unwrap_or(name);
            if let Ok(code) = u32::from_str_radix(hex_part, 16) {
                if let Some(ch) = char::from_u32(code) {
                    return ch.to_string();
                }
            }
            return format!("[{}]", name);
        }
    };
    result.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_strip_simple() {
        assert_eq!(strip_html("<p>Hello</p>"), "Hello");
    }

    #[test]
    fn test_strip_entities() {
        assert_eq!(strip_html("A &amp; B"), "A & B");
        assert_eq!(strip_html("&lt;code&gt;"), "<code>");
    }

    #[test]
    fn test_strip_nested() {
        assert_eq!(
            strip_html("<p>Hi <strong>team</strong>, how are you?</p>"),
            "Hi team, how are you?"
        );
    }

    #[test]
    fn test_strip_nbsp() {
        assert_eq!(strip_html("<p>&nbsp;</p>"), "");
    }

    #[test]
    fn test_plain_text_passthrough() {
        assert_eq!(strip_html("no html here"), "no html here");
    }

    #[test]
    fn test_blockquote_reply() {
        let html = r#"<div><blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="123"><p><b>Alice</b></p><p itemprop="preview">Hello world</p></blockquote>My reply</div>"#;
        let result = strip_html(html);
        assert!(result.contains("> Hello world"), "got: {}", result);
        assert!(result.contains("My reply"), "got: {}", result);
    }

    #[test]
    fn test_mention_prefix() {
        let html = r#"<p>Hey <span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">Dave</span>&nbsp;check this</p>"#;
        let result = strip_html(html);
        assert!(result.contains("@Dave"), "got: {}", result);
    }

    #[test]
    fn test_mention_multiword_name() {
        // Teams splits multi-word names into separate spans with same itemid
        let html = r#"<p>Hi <span itemtype="http://schema.skype.com/Mention" itemid="0">Jane</span> <span itemtype="http://schema.skype.com/Mention" itemid="0">Marie</span> <span itemtype="http://schema.skype.com/Mention" itemid="0">Smith</span>, check this</p>"#;
        let result = strip_html(html);
        assert!(result.contains("@Jane"), "should have mention: {}", result);
        // Same itemid = same person, only one @ prefix
        assert_eq!(result.matches('@').count(), 1, "should have exactly one @ prefix: {}", result);
    }

    #[test]
    fn test_mention_multiple_people() {
        // Two different people with different itemids
        let html = r#"<p>Hey <span itemtype="http://schema.skype.com/Mention" itemid="0">Alice</span> and <span itemtype="http://schema.skype.com/Mention" itemid="1">Bob</span></p>"#;
        let result = strip_html(html);
        assert!(result.contains("@Alice"), "got: {}", result);
        assert!(result.contains("@Bob"), "got: {}", result);
    }

    #[test]
    fn test_mention_different_ids_per_word() {
        // Teams gives each word of a name a different itemid
        // Adjacent mention spans (only whitespace between) should be joined with -
        let html = r#"<p>congratulating <span itemtype="http://schema.skype.com/Mention" itemid="1">John</span> <span itemtype="http://schema.skype.com/Mention" itemid="2">Doe</span> and <span itemtype="http://schema.skype.com/Mention" itemid="3">Jane</span> <span itemtype="http://schema.skype.com/Mention" itemid="4">Smith</span> on their work</p>"#;
        let result = strip_html(html);
        assert!(result.contains("@John-Doe"), "should join adjacent names: {}", result);
        assert!(result.contains("@Jane-Smith"), "should join adjacent names: {}", result);
        assert_eq!(result.matches('@').count(), 2, "should have exactly 2 mentions: {}", result);
    }

    #[test]
    fn test_rich_bold_italic() {
        let html = r#"<p>Hello <b>bold</b> and <i>italic</i> text</p>"#;
        let lines = strip_html_rich(html);
        let segments: Vec<&RichSegment> = lines.iter().flat_map(|l| l.iter()).collect();
        assert!(segments.iter().any(|s| s.text.contains("bold") && s.bold), "bold segment missing");
        assert!(segments.iter().any(|s| s.text.contains("italic") && s.italic), "italic segment missing");
    }

    #[test]
    fn test_rich_link() {
        let html = r#"<p>Check <a href="https://example.com">this link</a></p>"#;
        let lines = strip_html_rich(html);
        let segments: Vec<&RichSegment> = lines.iter().flat_map(|l| l.iter()).collect();
        assert!(segments.iter().any(|s| s.text.contains("this link") && s.link_url.is_some()), "link segment missing");
    }

    #[test]
    fn test_rich_mention_multiword() {
        // Multi-word name: all parts should be mention-styled, no duplicate @
        let html = r#"<p>Hi <span itemtype="http://schema.skype.com/Mention" itemid="0">Alice</span> <span itemtype="http://schema.skype.com/Mention" itemid="0">Wonder</span>, hello</p>"#;
        let lines = strip_html_rich(html);
        let segments: Vec<&RichSegment> = lines.iter().flat_map(|l| l.iter()).collect();
        let mention_segs: Vec<_> = segments.iter().filter(|s| s.mention).collect();
        assert!(mention_segs.len() >= 1, "should have mention segments, got: {:?}", segments);
        let mention_text: String = mention_segs.iter().map(|s| s.text.as_str()).collect();
        assert!(mention_text.contains("@Alice"), "should have @Alice: {}", mention_text);
        assert!(mention_text.contains("Wonder"), "should have Wonder: {}", mention_text);
        // Should NOT have @Wonder (no duplicate @)
        assert!(!mention_text.contains("@Wonder"), "should NOT have @Wonder: {}", mention_text);
    }

    #[test]
    fn test_rich_mention_spacing() {
        // Spaces around mention should be preserved
        let html = r#"<p>Hi <span itemtype="http://schema.skype.com/Mention" itemid="0">Carol</span> check this</p>"#;
        let lines = strip_html_rich(html);
        let segments: Vec<&RichSegment> = lines.iter().flat_map(|l| l.iter()).collect();
        // Reconstruct text and verify spacing
        let full: String = segments.iter().map(|s| s.text.as_str()).collect();
        assert!(full.contains("Hi "), "space before mention missing: '{}'", full);
        assert!(full.contains(" check"), "space after mention missing: '{}'", full);
    }

    #[test]
    fn test_mention_no_space_before() {
        // Mention directly after text (no space) should get a space inserted
        // Adjacent mention spans should be joined with -
        let html = r#"<p>results<span itemtype="http://schema.skype.com/Mention" itemid="0">Bob</span> <span itemtype="http://schema.skype.com/Mention" itemid="1">Jones</span></p>"#;
        let result = strip_html(html);
        assert!(result.contains("results @Bob-Jones"), "should join and space: {}", result);
    }

    #[test]
    fn test_rich_mention_adjacent_different_ids() {
        // Two adjacent mention spans with different itemids (same person, multi-word name)
        // Both words should be mention-highlighted and joined with -
        let html = r#"<p>thanks <span itemtype="http://schema.skype.com/Mention" itemid="0">Eve</span> <span itemtype="http://schema.skype.com/Mention" itemid="1">Parker</span>.</p>"#;
        let lines = strip_html_rich(html);
        let segments: Vec<&RichSegment> = lines.iter().flat_map(|l| l.iter()).collect();
        let mention_segs: Vec<_> = segments.iter().filter(|s| s.mention).collect();
        let mention_text: String = mention_segs.iter().map(|s| s.text.as_str()).collect();
        assert!(mention_text.contains("@Eve-"), "should have @Eve-: {}", mention_text);
        assert!(mention_text.contains("Parker"), "should have Parker highlighted: {}", mention_text);
    }

    #[test]
    fn test_stareyes_emoji() {
        assert_eq!(teams_emoji("stareyes"), "\u{1f929}");
    }
}
