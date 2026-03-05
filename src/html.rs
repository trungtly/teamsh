/// Strip HTML tags and decode common entities to plain text.
/// Handles Teams quote replies (blockquote) by prefixing with "> ".
pub fn strip_html(html: &str) -> String {
    // First, extract blockquote content and the reply separately
    let (quote, reply) = extract_quote(html);

    let stripped = strip_tags(&if let Some(reply) = &reply {
        reply.clone()
    } else {
        html.to_string()
    });

    if let Some(q) = quote {
        let q_text = strip_tags(&q);
        if !q_text.is_empty() {
            return format!("> {}\n{}", q_text, stripped);
        }
    }
    stripped
}

/// Extract (quote_preview, reply_text) from Teams HTML with blockquote.
fn extract_quote(html: &str) -> (Option<String>, Option<String>) {
    // Find blockquote boundaries
    let bq_start = html.find("<blockquote");
    let bq_end = html.find("</blockquote>");
    if let (Some(start), Some(end)) = (bq_start, bq_end) {
        let bq_content = &html[start..end];
        // Extract the preview text from itemprop="preview"
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

        // Reply is everything after </blockquote>
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

/// Strip HTML tags and entities from a string, collapse whitespace.
/// Prefixes @mentions with "@" based on schema.skype.com/Mention spans.
fn strip_tags(html: &str) -> String {
    let mut result = String::with_capacity(html.len());
    let mut in_tag = false;
    let mut in_entity = false;
    let mut entity = String::new();
    let mut tag_buf = String::new();
    let mut pending_mention = false;

    for ch in html.chars() {
        match ch {
            '<' => {
                in_tag = true;
                tag_buf.clear();
            }
            '>' if in_tag => {
                in_tag = false;
                // Check if this is a mention span opening tag
                if tag_buf.contains("schema.skype.com/Mention") && !tag_buf.starts_with('/') {
                    pending_mention = true;
                }
            }
            '&' if !in_tag => {
                if pending_mention {
                    result.push('@');
                    pending_mention = false;
                }
                in_entity = true;
                entity.clear();
            }
            ';' if in_entity => {
                in_entity = false;
                match entity.as_str() {
                    "amp" => result.push('&'),
                    "lt" => result.push('<'),
                    "gt" => result.push('>'),
                    "quot" => result.push('"'),
                    "nbsp" => result.push(' '),
                    "apos" => result.push('\''),
                    _ => {
                        result.push('&');
                        result.push_str(&entity);
                        result.push(';');
                    }
                }
            }
            _ if in_tag => {
                tag_buf.push(ch);
            }
            _ if in_entity => {
                entity.push(ch);
            }
            _ => {
                if pending_mention {
                    result.push('@');
                    pending_mention = false;
                }
                result.push(ch);
            }
        }
    }

    // Collapse multiple whitespace/newlines per line
    let mut collapsed = String::with_capacity(result.len());
    let mut last_was_space = false;
    for ch in result.trim().chars() {
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

/// Quick tag stripping for search indexing (no quote handling)
pub fn strip_tags_only(html: &str) -> String {
    strip_tags(html)
}

/// Map Teams emoji key to Unicode
pub fn teams_emoji(key: &str) -> String {
    // Strip blob/sticker reference after semicolon
    let base = key.split(';').next().unwrap_or(key);
    // Strip skin tone suffix (e.g. "-tone1", "-tone3")
    let name = if let Some(pos) = base.rfind("-tone") {
        &base[..pos]
    } else {
        base
    };

    let result = match name {
        // Standard Teams reactions
        "like" | "1f44d" | "1f44d_thumbsup" | "yes" => "\u{1f44d}",
        "heart" | "2764" | "2764_heart" => "\u{2764}",
        "laugh" | "1f606" | "1f606_laugh" => "\u{1f606}",
        "surprised" | "1f62e" | "1f62e_surprised" => "\u{1f62e}",
        "sad" | "1f622" | "1f622_sad" => "\u{1f622}",
        "angry" | "1f620" | "1f620_angry" => "\u{1f620}",
        // Common emoji
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
            // Try to parse hex emoji code from the key
            let hex_part = name.split('_').next().unwrap_or(name);
            if let Ok(code) = u32::from_str_radix(hex_part, 16) {
                if let Some(ch) = char::from_u32(code) {
                    return ch.to_string();
                }
            }
            // Show bracketed name for unknown custom stickers
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
        let html = r#"<p>Hey <span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">Ahmed</span>&nbsp;check this</p>"#;
        let result = strip_html(html);
        assert!(result.contains("@Ahmed"), "got: {}", result);
    }

    #[test]
    fn test_stareyes_emoji() {
        assert_eq!(teams_emoji("stareyes"), "\u{1f929}");
    }
}
