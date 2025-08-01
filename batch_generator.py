#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from chinese_qa_generator import ChineseQAGenerator
import os
import time

def batch_generate_qa(filename: str = "chinese_qa_data100000.xlsx", total_count: int = 100000, batch_size: int = 1000):
    """Generate Q&A pairs in batches with progress tracking."""
    
    print(f"Starting batch generation of {total_count} Q&A pairs...")
    print(f"Batch size: {batch_size}")
    print(f"Target file: {filename}")
    print("=" * 60)
    
    generator = ChineseQAGenerator()
    
    # Check if file exists
    file_exists = os.path.exists(filename)
    if file_exists:
        print(f"Found existing file '{filename}'. Will append new Q&A pairs.")
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
        
        batch_num += 1
        
        # Progress statistics
        elapsed_time = time.time() - start_time
        avg_time_per_qa = elapsed_time / total_generated
        remaining_qa = total_count - total_generated
        estimated_remaining_time = remaining_qa * avg_time_per_qa
        
        print(f"Average time per Q&A: {avg_time_per_qa:.3f} seconds")
        print(f"Estimated remaining time: {estimated_remaining_time/60:.1f} minutes")
    
    total_time = time.time() - start_time
    print(f"\n" + "=" * 60)
    print(f"BATCH GENERATION COMPLETED!")
    print(f"Total Q&A pairs generated: {total_generated}")
    print(f"Total time: {total_time/60:.1f} minutes")
    print(f"Average time per Q&A: {total_time/total_generated:.3f} seconds")
    print(f"File saved as: {filename}")
    
    return total_generated

def main():
    """Main function for batch generation."""
    print("Enhanced Chinese Q&A Batch Generator")
    print("=" * 60)
    
    # Configuration
    filename = "chinese_qa_data100000.xlsx"
    total_count = 100000
    batch_size = 1000  # Process in batches of 1000
    
    print(f"Configuration:")
    print(f"- Target file: {filename}")
    print(f"- Total Q&A pairs: {total_count}")
    print(f"- Batch size: {batch_size}")
    print(f"- Estimated batches: {total_count // batch_size + (1 if total_count % batch_size else 0)}")
    
    # Start generation
    batch_generate_qa(filename, total_count, batch_size)

if __name__ == "__main__":
    main() 