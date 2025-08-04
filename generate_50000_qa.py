#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from chinese_qa_generator import ChineseQAGenerator
import os
import time

def generate_50000_qa(filename: str = "chinese_qa_50000.xlsx", batch_size: int = 1000):
    """Generate exactly 50,000 unique Q&A pairs with progress tracking."""
    
    total_count = 50000
    print(f"Starting generation of {total_count} unique Q&A pairs...")
    print(f"Target file: {filename}")
    print(f"Batch size: {batch_size}")
    print("=" * 60)
    
    generator = ChineseQAGenerator()
    
    # Check if file exists
    file_exists = os.path.exists(filename)
    if file_exists:
        print(f"Found existing file '{filename}'. Will append new Q&A pairs.")
        # Load existing questions to avoid duplicates
        existing_questions = generator.load_existing_questions(filename)
        generator.used_questions.update(existing_questions)
        print(f"Loaded {len(existing_questions)} existing questions to avoid duplicates.")
    else:
        print(f"Creating new file '{filename}'.")
    
    total_generated = 0
    batch_num = 1
    start_time = time.time()
    
    while total_generated < total_count:
        current_batch_size = min(batch_size, total_count - total_generated)
        
        print(f"\n--- Batch {batch_num} ---")
        print(f"Generating {current_batch_size} Q&A pairs...")
        
        batch_start_time = time.time()
        
        # Generate batch
        qa_pairs = generator.generate_qa_pairs(current_batch_size)
        
        # Write to Excel
        if batch_num == 1 and not file_exists:
            generator.write_to_excel(qa_pairs, filename, append=False)
        else:
            generator.write_to_excel(qa_pairs, filename, append=True)
        
        batch_time = time.time() - batch_start_time
        total_generated += len(qa_pairs)
        
        print(f"Batch {batch_num} completed in {batch_time:.2f} seconds")
        print(f"Generated: {len(qa_pairs)} Q&A pairs")
        print(f"Total progress: {total_generated}/{total_count} ({total_generated/total_count*100:.1f}%)")
        
        # Progress statistics
        elapsed_time = time.time() - start_time
        avg_time_per_qa = elapsed_time / total_generated
        remaining_qa = total_count - total_generated
        estimated_remaining_time = remaining_qa * avg_time_per_qa
        
        print(f"Average time per Q&A: {avg_time_per_qa:.3f} seconds")
        print(f"Estimated remaining time: {estimated_remaining_time/60:.1f} minutes")
        
        batch_num += 1
    
    total_time = time.time() - start_time
    print(f"\n" + "=" * 60)
    print(f"GENERATION COMPLETED!")
    print(f"Total Q&A pairs generated: {total_generated}")
    print(f"Total time: {total_time/60:.1f} minutes")
    print(f"Average time per Q&A: {total_time/total_generated:.3f} seconds")
    print(f"File saved as: {filename}")
    
    # Verify uniqueness
    print(f"\nVerifying uniqueness...")
    generator.load_existing_questions(filename)
    print(f"Total unique questions in file: {len(generator.used_questions)}")
    
    return total_generated

def main():
    """Main function for 50,000 Q&A generation."""
    print("50,000 Chinese Q&A Generator")
    print("=" * 60)
    
    # Configuration
    filename = "chinese_qa_50000.xlsx"
    batch_size = 1000  # Process in batches of 1000
    
    print(f"Configuration:")
    print(f"- Target file: {filename}")
    print(f"- Total Q&A pairs: 50,000")
    print(f"- Batch size: {batch_size}")
    print(f"- Estimated batches: {50000 // batch_size + (1 if 50000 % batch_size else 0)}")
    print(f"- Question templates available: 49")
    print(f"- Topic categories: 8")
    print(f"- Total topics available: 125")
    print(f"- Answer patterns available: 23")
    
    # Start generation
    generate_50000_qa(filename, batch_size)

if __name__ == "__main__":
    main() 