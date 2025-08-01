#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from chinese_qa_generator import ChineseQAGenerator
import os

def append_qa_pairs(filename: str = "chinese_qa_data.xlsx", count: int = 20):
    """Append new Q&A pairs to existing Excel file."""
    
    if not os.path.exists(filename):
        print(f"File '{filename}' does not exist. Creating new file...")
        generator = ChineseQAGenerator()
        generator.generate_and_save(count=count, filename=filename, append=False)
        return
    
    print(f"Appending {count} new Q&A pairs to existing file '{filename}'...")
    
    generator = ChineseQAGenerator()
    qa_pairs = generator.generate_and_save(count=count, filename=filename, append=True)
    
    if qa_pairs:
        print(f"\nSuccessfully added {len(qa_pairs)} new Q&A pairs!")
        print("\nNew Q&A pairs added:")
        print("-" * 50)
        
        for i, qa in enumerate(qa_pairs, 1):
            print(f"\n{i}. 问题: {qa['标准问题']}")
            print(f"   回答类型: {qa['回答类型']}")
            print(f"   回答: {qa['问题回答1']}")
            print(f"   字符数: {len(qa['问题回答1'])}")
    else:
        print("No new Q&A pairs were added (all questions already existed).")

def main():
    """Main function to append Q&A pairs."""
    print("Chinese Q&A Appender")
    print("=" * 50)
    
    # You can modify these parameters
    filename = "chinese_qa_data100000.xlsx"
    count = 100000  # Number of new Q&A pairs to add
    
    append_qa_pairs(filename, count)

if __name__ == "__main__":
    main() 