#!/usr/bin/env python3
"""
Enhanced DOC to Excel Converter - Based on original code

This is the enhanced version of the original converter with improved
functionality and structure for GitHub repository use.

Original Author: i2213
Enhanced Version: 1.0.0
"""

import os
import re
from typing import List, Dict, Tuple, Optional, Union

import pandas as pd
from docx import Document
from docx.text.run import Run
from docx.enum.text import WD_COLOR_INDEX

# Enhanced Regex patterns for medical questions
QUESTION_HEADER_RE = re.compile(r"^\s*Question\s+(\d+)(?:\s*[-–—]\s*(.+))?\s*$", re.I)
OPTION_RE = re.compile(r"^([A-E])[\.\)]\s*(.*)", re.S)
CORRECT_RE_LIST = [
    re.compile(r"^\s*Correct\s*Answer\s*[:\-]?\s*([A-Ea-e])\s*$", re.I),
    re.compile(r"^\s*Answer\s*[:\-]?\s*([A-Ea-e])\s*$", re.I),
]
EXPLANATION_HEADER_RE = re.compile(r"^\s*Explanation(\s+(of\s+the\s+)?Correct\s+Answer)?\s*[:\-]?\s*$", re.I)
SECTION_HEADERS = [
    re.compile(r"^\s*Analysis\s+of\s+Other\s+Options", re.I),
    re.compile(r"^\s*Analysis\s+of\s+Distractors", re.I),
    re.compile(r"^\s*Distractors", re.I),
    re.compile(r"^\s*Key\s+Insights?", re.I),
    re.compile(r"^\s*Question\s+\d+", re.I),
]

# Answer mapping
LETTER_TO_NUM = {"A": "1", "B": "2", "C": "3", "D": "4", "E": "5"}
NUM_TO_LETTER = {v: k for k, v in LETTER_TO_NUM.items()}


def normalize_text(text: str) -> str:
    """Normalize whitespace and clean text."""
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\u00a0", " ", text)              # non-breaking space
    text = re.sub(r"[\u2018\u2019]", "'", text)      # smart single quotes
    text = re.sub(r"[\u201c\u201d]", '"', text)      # smart double quotes
    text = re.sub(r"[\u2013\u2014]", "-", text)      # en/em dashes to hyphen
    return text.strip()


def replace_option_letters_with_numbers(text: str) -> str:
    """Convert A/B/C/D/E references into numbered format 1-/2-/3-/4-/5- in explanations."""
    if not text:
        return text
    # Replace standalone A/B/C/D/E with number-
    text = re.sub(r"\bA[\)\.]", "1-", text)
    text = re.sub(r"\bB[\)\.]", "2-", text)
    text = re.sub(r"\bC[\)\.]", "3-", text)
    text = re.sub(r"\bD[\)\.]", "4-", text)
    text = re.sub(r"\bE[\)\.]", "5-", text)
    return text


def is_highlighted(run) -> bool:
    """Check if a text run is highlighted."""
    try:
        if hasattr(run.font, "highlight_color") and run.font.highlight_color is not None:
            return True
        if hasattr(run, "_element"):
            highlight = run._element.xpath(".//w:highlight")
            if highlight:
                return True
    except Exception:
        pass
    return False


def extract_highlighted_text(paragraph) -> Tuple[str, str]:
    """Extract highlighted and non-highlighted text from a paragraph."""
    highlighted_parts, normal_parts = [], []
    for run in paragraph.runs:
        if not run.text:
            continue
        if is_highlighted(run):
            highlighted_parts.append(run.text)
        else:
            normal_parts.append(run.text)
    highlighted_text = normalize_text(" ".join(highlighted_parts))
    normal_text = normalize_text(" ".join(normal_parts))
    if not highlighted_text:
        return "", normalize_text(paragraph.text)
    return highlighted_text, normal_text


def find_question_blocks(paragraphs) -> List[Tuple[int, int, str, str]]:
    """Find question blocks with their numbers and topics."""
    blocks = []
    for i, para in enumerate(paragraphs):
        text = normalize_text(para.text)
        m = QUESTION_HEADER_RE.match(text)
        if m:
            question_num = m.group(1)
            topic = m.group(2) if m.group(2) else ""
            # End of block is next question header or end of doc
            end_idx = len(paragraphs)
            for j in range(i + 1, len(paragraphs)):
                next_text = normalize_text(paragraphs[j].text)
                if QUESTION_HEADER_RE.match(next_text):
                    end_idx = j
                    break
            blocks.append((i, end_idx, question_num, topic))
    return blocks


def parse_options_section(paragraphs, start_idx: int, end_idx: int) -> Tuple[Dict[str, str], int]:
    """Parse options and return options dict and index where options end."""
    options: Dict[str, str] = {}
    current_option: Optional[str] = None
    current_text_parts: List[str] = []
    options_end_idx = end_idx

    def save_current_option():
        nonlocal options, current_option, current_text_parts
        if current_option and current_text_parts:
            options[current_option] = normalize_text(" ".join(current_text_parts))

    for i in range(start_idx, end_idx):
        text = normalize_text(paragraphs[i].text)
        if not text:
            continue

        # Stop at section boundary, answer line, or explanation header
        if any(patt.match(text) for patt in SECTION_HEADERS):
            options_end_idx = i
            break
        if any(patt.match(text) for patt in CORRECT_RE_LIST):
            options_end_idx = i
            break
        if EXPLANATION_HEADER_RE.match(text):
            options_end_idx = i
            break

        # Option line
        option_match = OPTION_RE.match(text)
        if option_match:
            save_current_option()
            current_option = option_match.group(1).upper()
            option_text = option_match.group(2).strip()
            current_text_parts = [option_text] if option_text else []
        else:
            # Continuation line for the current option
            if current_option:
                current_text_parts.append(text)

    save_current_option()
    return options, options_end_idx


def find_correct_answer(paragraphs, start_idx: int, end_idx: int) -> Tuple[Optional[str], int]:
    """Find the correct answer and return it with the index after the answer line."""
    for i in range(start_idx, end_idx):
        text = normalize_text(paragraphs[i].text)
        for patt in CORRECT_RE_LIST:
            match = patt.match(text)
            if match:
                return match.group(1).upper(), i + 1
    return None, start_idx


def parse_explanation_section(paragraphs, start_idx: int, end_idx: int) -> Tuple[str, str]:
    """Parse explanation section, returning (correct_explanation, incorrect_explanation)."""
    highlighted_parts: List[str] = []
    normal_parts: List[str] = []

    # Find "Explanation" header
    explanation_start = start_idx
    for i in range(start_idx, end_idx):
        text = normalize_text(paragraphs[i].text)
        if EXPLANATION_HEADER_RE.match(text):
            explanation_start = i + 1
            break

    # Explanation ends at next section header or end
    explanation_end = end_idx
    for i in range(explanation_start, end_idx):
        text = normalize_text(paragraphs[i].text)
        if any(patt.match(text) for patt in SECTION_HEADERS):
            explanation_end = i
            break

    # Collect highlighted vs normal
    for i in range(explanation_start, explanation_end):
        hi, norm = extract_highlighted_text(paragraphs[i])
        if hi:
            highlighted_parts.append(hi)
        if norm:
            normal_parts.append(norm)

    highlighted_text = " ".join(highlighted_parts).strip()
    normal_text = " ".join(normal_parts).strip()

    # If no highlighting, treat all text as the correct explanation
    if not highlighted_text and normal_text:
        highlighted_text = normal_text
        normal_text = ""

    # Apply option letter to number conversion in explanations
    highlighted_text = replace_option_letters_with_numbers(highlighted_text)
    normal_text = replace_option_letters_with_numbers(normal_text)

    return highlighted_text, normal_text


def parse_question_stem(paragraphs, start_idx: int, options_start_idx: int) -> str:
    """Parse the question stem from paragraphs."""
    stem_parts: List[str] = []
    for i in range(start_idx, options_start_idx):
        text = normalize_text(paragraphs[i].text)
        if text:
            stem_parts.append(text)
    return " ".join(stem_parts)


def parse_single_question(
    paragraphs, start_idx: int, end_idx: int, question_num: str, topic: str, source_file: str, 
    group: str = "Medical", qtype: str = "MCQ"
) -> Dict[str, str]:
    """Parse a single question block."""

    # Skip the header line
    content_start = start_idx + 1

    # Find the first option line A./B./C./D./E.
    options_start_idx = content_start
    for i in range(content_start, end_idx):
        text = normalize_text(paragraphs[i].text)
        if OPTION_RE.match(text):
            options_start_idx = i
            break

    # Parse stem
    question_stem = parse_question_stem(paragraphs, content_start, options_start_idx)

    # Parse options
    options, options_end_idx = parse_options_section(paragraphs, options_start_idx, end_idx)

    # Correct answer
    correct_answer, answer_end_idx = find_correct_answer(paragraphs, options_end_idx, end_idx)

    # Explanations
    explanation_start_idx = max(options_end_idx, answer_end_idx)
    correct_explanation, incorrect_explanation = parse_explanation_section(
        paragraphs, explanation_start_idx, end_idx
    )

    # Fallback: if incorrect explanation missing, copy correct explanation
    if not (incorrect_explanation or "").strip():
        incorrect_explanation = correct_explanation

    result = {
        "MCQ_NO": question_num,
        "Question": question_stem,
        "Answer1": options.get("A", ""),
        "Answer2": options.get("B", ""),
        "Answer3": options.get("C", ""),
        "Answer4": options.get("D", ""),
        "Answer5": options.get("E", ""),
        "CorAns": LETTER_TO_NUM.get(correct_answer, "") if correct_answer else "",
        "CorrectExplanation": correct_explanation,
        "IncorrectExplanation": incorrect_explanation,
        "Topic": topic,
        "SourceFile": os.path.basename(source_file),
        "Group": group,
        "Type": qtype,
    }
    return result


def parse_docx_file(file_path: str, group: str = "Medical", qtype: str = "MCQ") -> List[Dict[str, str]]:
    """Parse a single DOCX file and extract all questions."""
    try:
        doc = Document(file_path)
        paragraphs = doc.paragraphs

        # Find question blocks
        question_blocks = find_question_blocks(paragraphs)
        if not question_blocks:
            print(f"Warning: No questions found in {file_path}")
            return []

        results: List[Dict[str, str]] = []
        for start_idx, end_idx, question_num, topic in question_blocks:
            try:
                question_data = parse_single_question(
                    paragraphs, start_idx, end_idx, question_num, topic, file_path, group, qtype
                )
                # Must have a question and at least one option
                if question_data["Question"] and any(
                    question_data[f"Answer{i}"] for i in range(1, 6)
                ):
                    results.append(question_data)
                else:
                    print(f"Warning: Incomplete question {question_num} in {file_path}")
            except Exception as e:
                print(f"Error parsing question {question_num} in {file_path}: {e}")
                continue

        print(f"Extracted {len(results)} questions from {file_path}")
        return results

    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return []


def process_files_to_excel(
    input_path: str, output_path: str, group: str = "Medical", qtype: str = "MCQ"
) -> Optional[pd.DataFrame]:
    """Process Word files and convert to Excel format."""

    # Determine input files
    if os.path.isfile(input_path):
        if not input_path.lower().endswith(".docx"):
            print(f"Error: {input_path} is not a .docx file")
            return None
        input_files = [input_path]
    elif os.path.isdir(input_path):
        input_files = [
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".docx") and not f.startswith("~")
        ]
        if not input_files:
            print(f"Error: No .docx files found in {input_path}")
            return None
    else:
        print(f"Error: {input_path} does not exist")
        return None

    # Process all files
    all_questions: List[Dict[str, str]] = []
    processed_files = 0
    
    print(f"🚀 Processing {len(input_files)} files...")
    
    for file_path in input_files:
        print(f"\n📄 Processing: {os.path.basename(file_path)}")
        questions = parse_docx_file(file_path, group, qtype)
        if questions:
            all_questions.extend(questions)
            processed_files += 1

    if not all_questions:
        print("No questions extracted from any files.")
        return None

    # Create DataFrame
    columns = [
        "MCQ_NO",
        "Question",
        "Answer1",
        "CorAns",
        "Answer2",
        "Answer3",
        "Answer4",
        "Answer5",
        "CorrectExplanation",
        "IncorrectExplanation",
        "Topic",
        "Group",
        "Type",
        "SourceFile",
    ]
    df = pd.DataFrame(all_questions, columns=columns)

    # Clean text
    text_columns = [
        "Question",
        "Answer1",
        "Answer2",
        "Answer3",
        "Answer4",
        "Answer5",
        "CorrectExplanation",
        "IncorrectExplanation",
        "Topic",
    ]
    for col in text_columns:
        df[col] = df[col].astype(str).apply(normalize_text)

    # Safety net: fill empty IncorrectExplanation with CorrectExplanation
    if "CorrectExplanation" in df.columns and "IncorrectExplanation" in df.columns:
        mask_empty_ie = df["IncorrectExplanation"].astype(str).str.strip().eq("")
        df.loc[mask_empty_ie, "IncorrectExplanation"] = df.loc[mask_empty_ie, "CorrectExplanation"]

    # Save to Excel
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Questions", index=False)

            # Create summary sheet
            import datetime
            summary_df = pd.DataFrame(
                [
                    ["Processing Summary", ""],
                    ["Total Questions", len(df)],
                    ["Files Processed", processed_files],
                    ["Questions per File", f"{len(df) / max(1, processed_files):.1f}"],
                    ["Processing Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    ["", ""],
                    ["Question Distribution by Group", ""],
                    [f"  {group}", len(df)],
                    ["", ""],
                    ["Question Distribution by Type", ""],
                    [f"  {qtype}", len(df)],
                ],
                columns=["Metric", "Value"],
            )
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

        print(f"\n✅ Successfully created Excel file with {len(df)} questions: {output_path}")
        return df

    except Exception as e:
        print(f"Error saving Excel file: {e}")
        return None


def main():
    """Main function for command-line usage."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Enhanced DOC to Excel Converter for MCQ Processing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Convert single file
    python enhanced_converter.py --input "questions.docx" --output "output.xlsx"
    
    # Convert directory with custom settings
    python enhanced_converter.py --input "docs/" --output "medical_questions.xlsx" \\
                                --group "Internal Medicine" --type "Board Review"
        """
    )
    
    parser.add_argument("--input", "-i", required=True, help="Input file or directory")
    parser.add_argument("--output", "-o", required=True, help="Output Excel file")
    parser.add_argument("--group", "-g", default="Medical", help="Question group (default: Medical)")
    parser.add_argument("--type", "-t", default="MCQ", help="Question type (default: MCQ)")
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("ENHANCED DOC to EXCEL CONVERTER")
    print("=" * 60)
    print(f"📂 Input: {args.input}")
    print(f"📊 Output: {args.output}")
    print(f"🏷️  Group: {args.group}")
    print(f"📝 Type: {args.type}")
    print()
    
    result = process_files_to_excel(args.input, args.output, args.group, args.type)
    
    if result is not None:
        print(f"\n🎉 Processing completed successfully!")
        print(f"📈 Total questions processed: {len(result)}")
        print(f"📄 Output file: {args.output}")
    else:
        print("\n❌ Processing failed. Please check the error messages above.")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())