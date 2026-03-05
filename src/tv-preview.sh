#!/bin/bash
# tv preview script for teamsh messages
# Input: rg output line like "/path/to/messages/1234.txt:5:matched content"

entry="$1"

# Extract file path (everything before .txt: + .txt)
file=""
if [[ "$entry" =~ ^(.+\.txt):[0-9]+: ]]; then
    file="${BASH_REMATCH[1]}"
fi
[ -z "$file" ] && exit 0

dir=$(dirname "$file")
base=$(basename "$dir")

if [ "$base" = "messages" ]; then
    # Message file - show all messages in conversation with separators
    conv_dir=$(dirname "$dir")
    conv_name=""
    if [ -f "$conv_dir/meta.json" ]; then
        conv_name=$(grep -o '"name":"[^"]*"' "$conv_dir/meta.json" 2>/dev/null | head -1 | sed 's/"name":"//;s/"$//')
    fi

    # Find which line in the matched file
    line_in_file=""
    if [[ "$entry" =~ \.txt:([0-9]+): ]]; then
        line_in_file="${BASH_REMATCH[1]}"
    fi

    # Build concatenated output and track line offset for highlight
    highlight_line=0
    current_line=0
    matched_file=$(basename "$file")
    tmp=$(mktemp)

    if [ -n "$conv_name" ]; then
        echo "$conv_name" >> "$tmp"
        echo "────────────────────────────────" >> "$tmp"
        current_line=2
    fi

    first=1
    for f in $(ls "$dir"/*.txt 2>/dev/null | sort); do
        fname=$(basename "$f")
        if [ "$first" = "1" ]; then
            first=0
        else
            echo "" >> "$tmp"
            current_line=$((current_line + 1))
        fi

        content=$(cat "$f" 2>/dev/null)
        line_count=$(echo "$content" | wc -l)

        if [ "$fname" = "$matched_file" ] && [ -n "$line_in_file" ]; then
            highlight_line=$((current_line + line_in_file))
        fi

        echo "$content" >> "$tmp"
        current_line=$((current_line + line_count))
    done

    if [ "$highlight_line" -gt 0 ]; then
        bat --language=teamsmsg --style=plain --color=always --paging=never --theme=teamsh --highlight-line="$highlight_line" "$tmp"
    else
        bat --language=teamsmsg --style=plain --color=always --paging=never --theme=teamsh "$tmp"
    fi
    rm -f "$tmp"
else
    # Regular file - just show it with bat
    bat --language=teamsmsg --style=plain --color=always --paging=never --theme=teamsh "$file" 2>/dev/null
fi
