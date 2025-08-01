# Chinese Q&A Generator

A Python tool that generates random Chinese questions and answers, then writes them to an Excel table.

## Features

- **Enhanced Diversity**: 50+ question templates across 6 categories (basic, specific, comparative, process, problem-solving, future-oriented)
- **Rich Topic Categories**: 120+ specific topics across 8 categories (AI/ML, BigData, Cloud, DevOps, Security, Database, Mobile, Web)
- **Complex Answer Patterns**: 20+ sophisticated answer patterns with context-aware component selection
- **Duplicate Prevention**: Advanced duplicate detection with smart question generation
- **Batch Processing**: Efficient batch generation with progress tracking for large datasets
- **Excel Integration**: Creates properly formatted Excel files with green headers
- **Answer Types**: Supports both "纯文本" and "富文本" answer types
- **Character Limit**: Ensures all answers stay within 200 characters

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the script to generate 50 Q&A pairs by default:

```bash
python chinese_qa_generator.py
```

### Custom Usage

You can also use the generator in your own Python code:

```python
from chinese_qa_generator import ChineseQAGenerator

# Create generator instance
generator = ChineseQAGenerator()

# Generate custom number of Q&A pairs
qa_pairs = generator.generate_and_save(count=100, filename="my_qa_data.xlsx")

# Append new Q&A pairs to existing file (won't overwrite)
qa_pairs = generator.generate_and_save(count=50, filename="my_qa_data.xlsx", append=True)
```

### Appending to Existing Files

To add more Q&A pairs to an existing Excel file without overwriting:

```bash
python append_qa.py
```

This will add 100,000 new Q&A pairs to the existing `chinese_qa_data100000.xlsx` file.

### Batch Generation with Progress Tracking

For large-scale generation with progress tracking:

```bash
python batch_generator.py
```

This will generate 100,000 Q&A pairs in batches of 1,000 with real-time progress updates.

### Demo Enhanced Generator

To see the enhanced generator in action with sample output:

```bash
python demo_enhanced_generator.py
```

## Output Format

The tool generates an Excel file with three columns:

1. **标准问题 (必填)** - Standard Question (Required)
2. **回答类型 (必填)** - Answer Type (Required) - Either "纯文本" or "富文本"
3. **问题回答1 (必填)** - Question Answer 1 (Required)

## Example Output

| 标准问题 (必填) | 回答类型 (必填) | 问题回答1 (必填) |
|----------------|----------------|----------------|
| 什么是人工智能？ | 纯文本 | 人工智能是一种机器学习技术，主要用于大数据。 |
| 区块链的特点是什么？ | 富文本 | 区块链指的是分布式系统，具有网络安全的特点。 |

## Customization

You can modify the following in the `ChineseQAGenerator` class:

- `question_templates`: Add more question templates
- `topics`: Add more topics for questions
- `answer_patterns`: Add more answer generation patterns
- `answer_types`: Modify answer types

## Requirements

- Python 3.6+
- openpyxl
- pandas
- faker (for additional randomization if needed)

## License

This project is open source and available under the MIT License. 