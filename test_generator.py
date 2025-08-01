#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from chinese_qa_generator import ChineseQAGenerator

def test_generator():
    """Test the Chinese Q&A generator with a small sample."""
    print("Testing Chinese Q&A Generator...")
    print("=" * 50)
    
    # Create generator instance
    generator = ChineseQAGenerator()
    
    # Generate 10 Q&A pairs for testing
    print("Generating 10 Q&A pairs...")
    qa_pairs = generator.generate_qa_pairs(10)
    
    # Display the results
    print("\nGenerated Q&A pairs:")
    print("-" * 50)
    
    for i, qa in enumerate(qa_pairs, 1):
        print(f"\n{i}. 问题: {qa['标准问题']}")
        print(f"   回答类型: {qa['回答类型']}")
        print(f"   回答: {qa['问题回答1']}")
        print(f"   字符数: {len(qa['问题回答1'])}")
    
    # Test Excel generation
    print("\n" + "=" * 50)
    print("Testing Excel file generation...")
    
    filename = "test_qa_data.xlsx"
    generator.write_to_excel(qa_pairs, filename)
    
    print(f"Excel file '{filename}' created successfully!")
    print("You can open the file to see the formatted table.")

if __name__ == "__main__":
    test_generator() 