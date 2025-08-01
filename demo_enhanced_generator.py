#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from chinese_qa_generator import ChineseQAGenerator

def demo_enhanced_generator():
    """Demonstrate the enhanced Chinese Q&A generator capabilities."""
    print("Enhanced Chinese Q&A Generator Demo")
    print("=" * 60)
    
    generator = ChineseQAGenerator()
    
    # Generate 15 Q&A pairs to show variety
    print("Generating 15 diverse Q&A pairs...")
    qa_pairs = generator.generate_qa_pairs(15)
    
    print(f"\nGenerated {len(qa_pairs)} unique Q&A pairs!")
    print("\nSample Q&A pairs showing variety:")
    print("-" * 60)
    
    for i, qa in enumerate(qa_pairs, 1):
        print(f"\n{i:2d}. 问题: {qa['标准问题']}")
        print(f"    回答类型: {qa['回答类型']}")
        print(f"    回答: {qa['问题回答1']}")
        print(f"    字符数: {len(qa['问题回答1'])}")
        print("-" * 40)
    
    # Show statistics
    print(f"\nStatistics:")
    print(f"- Total unique questions generated: {len(generator.used_questions)}")
    print(f"- Question templates available: {len(generator.question_templates)}")
    print(f"- Topic categories: {len(generator.topics)}")
    print(f"- Total topics available: {len(generator.all_topics)}")
    print(f"- Answer patterns available: {len(generator.answer_patterns)}")
    
    # Test Excel generation
    print(f"\n" + "=" * 60)
    print("Testing Excel file generation...")
    
    filename = "enhanced_qa_demo.xlsx"
    generator.write_to_excel(qa_pairs, filename)
    
    print(f"Excel file '{filename}' created successfully!")
    print("You can open the file to see the formatted table with enhanced Q&A pairs.")

if __name__ == "__main__":
    demo_enhanced_generator() 