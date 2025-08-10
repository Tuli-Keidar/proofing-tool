#!/usr/bin/env python3
"""
Document Proofreader class - handles document processing and OpenAI integration
"""

import os
import json
import re
from docx import Document
import tiktoken
from openai import AzureOpenAI

from config import Config
from utils import calculate_costs


class DocumentProofreader:
    def __init__(self):
        """Initialize the document proofreader with Azure OpenAI client"""
        # Validate configuration first
        Config.validate()
        
        self.client = AzureOpenAI(
            api_key=Config.AZURE_OPENAI_API_KEY,
            api_version=Config.AZURE_API_VERSION,
            azure_endpoint=Config.AZURE_OPENAI_ENDPOINT
        )
        self.encoder = tiktoken.encoding_for_model("gpt-4")
        
    def extract_document_structure(self, docx_path):
        """
        Extract document structure using python-docx
        Returns a list of sections with hierarchical information
        """
        print(f"Opening document: {docx_path}")
        doc = Document(docx_path)
        print(f"Document opened successfully with {len(doc.paragraphs)} paragraphs")
        
        sections = []
        
        # Start with a default root section
        root_section = {
            "title": "Document Root",
            "content": [],
            "level": 0,
            "parent_idx": None,
            "section_id": 0,
            "children": []
        }
        sections.append(root_section)
        
        current_section = root_section
        section_stack = [root_section]  # Track section hierarchy
        
        # Function to determine heading level from style name
        def get_heading_level(paragraph):
            if not paragraph.style.name.startswith('Heading'):
                return None
            try:
                return int(paragraph.style.name.replace('Heading ', ''))
            except ValueError:
                return 1  # Default to level 1 if can't parse

        # Process all paragraphs
        for para_idx, para in enumerate(doc.paragraphs):
            # Skip empty paragraphs
            if not para.text.strip():
                continue
                
            heading_level = get_heading_level(para)
            
            if heading_level is not None:
                # This is a heading - create a new section
                
                # Adjust the section stack based on heading level
                while len(section_stack) > 1 and section_stack[-1]["level"] >= heading_level:
                    section_stack.pop()
                
                # Find parent in the section stack
                parent = section_stack[-1]  # Get the last item in the stack
                parent_idx = parent["section_id"]
                
                # Create new section
                new_section = {
                    "title": para.text,
                    "content": [],
                    "level": heading_level,
                    "parent_idx": parent_idx,
                    "section_id": len(sections),
                    "children": []
                }
                
                # Add this section as a child to its parent
                if parent_idx < len(sections):  # Ensure parent index is valid
                    sections[parent_idx]["children"].append(len(sections))
                
                # Add to sections list
                sections.append(new_section)
                
                # Update current section and stack
                current_section = new_section
                section_stack.append(new_section)
            else:
                # Regular paragraph - add to current section
                current_section["content"].append(para.text)
        
        # Calculate token counts for all sections
        for section in sections:
            section_text = "\n".join(section["content"])
            section["token_count"] = len(self.encoder.encode(section_text))
        
        # Log section information
        print(f"Extracted {len(sections)} sections:")
        for i, section in enumerate(sections):
            print(f"  Section {i}: '{section['title']}' - Level {section['level']} - {len(section['content'])} paragraphs")
        
        return sections
    
    def segment_document(self, sections, max_tokens=None):
        """Segment the document based on structural boundaries and token limits"""
        if max_tokens is None:
            max_tokens = Config.MAX_TOKENS_PER_SEGMENT
            
        # Skip the root section (if it exists)
        content_sections = sections[1:] if len(sections) > 1 else sections
        
        # If we have very few sections, just return them all as one segment
        if len(content_sections) <= 1:
            return [[0]] if len(sections) == 1 else [[0, 1]]
        
        # Extract section texts and calculate token counts
        token_counts = []
        
        for section in content_sections:
            token_counts.append(section["token_count"])
        
        # Handle empty sections
        if all(count == 0 for count in token_counts):
            print("Warning: No content found in sections. Returning all as one segment.")
            return [list(range(len(sections)))]
        
        print(f"Segmenting {len(content_sections)} content sections")
        
        # Create segments based on document structure and token limits
        segments = []
        current_segment = []
        current_tokens = 0
        
        # Start with the root section if it exists
        if len(sections) > 1:
            current_segment = [0]
        
        # Process each content section
        for i, section in enumerate(content_sections):
            # The actual section index in the full sections list
            section_idx = i + 1 if len(sections) > 1 else i
            section_token_count = token_counts[i]
            
            # If this would be the first section in the segment or we're under token limit
            if not current_segment or (current_tokens + section_token_count <= max_tokens):
                # Add to current segment
                current_segment.append(section_idx)
                current_tokens += section_token_count
            else:
                # This section would exceed the token limit
                # Save current segment and start a new one
                if current_segment:
                    segments.append(current_segment)
                current_segment = [section_idx]
                current_tokens = section_token_count
                
            # Special case: if this section alone exceeds the token limit,
            # we need to split it further in the future (just add warning for now)
            if section_token_count > max_tokens:
                print(f"Warning: Section {section_idx} ('{section['title']}') exceeds maximum token limit ({section_token_count} > {max_tokens})")
        
        # Add the last segment
        if current_segment:
            segments.append(current_segment)
            
        print(f"Created {len(segments)} segments")
        for i, segment in enumerate(segments):
            segment_token_count = sum(sections[idx]["token_count"] for idx in segment)
            print(f"  Segment {i}: {len(segment)} sections - {segment_token_count} tokens - {[sections[idx]['title'] for idx in segment]}")
            
        return segments
    
    def proofread_content(self, content):
        """Send content to GPT-4o for proofreading"""
        try:
            user_prompt = Config.USER_PROMPT_TEMPLATE.format(content=content)
            
            response = self.client.chat.completions.create(
                model=Config.MODEL_NAME,
                messages=[
                    {"role": "system", "content": Config.SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=4000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            print(f"Error proofreading content: {str(e)}")
            return f"Error during proofreading: {str(e)}"
    
    def create_docx_report(self, content, output_path, include_costs=False, input_tokens=0, output_tokens=0):
        """Create a nicely formatted DOCX report from markdown content"""
        doc = Document()
        
        # Add document title
        doc.add_heading('Document Proofreading Report', 0)
        
        # Add cost analysis if requested
        if include_costs:
            costs = calculate_costs(input_tokens, output_tokens)
            doc.add_heading('Cost Analysis', 1)
            doc.add_paragraph(f"• Total Input Tokens: {costs['input_tokens']:,}")
            doc.add_paragraph(f"• Total Output Tokens: {costs['output_tokens']:,}")
            doc.add_paragraph(f"• Estimated API Cost: ${costs['total_cost']:.4f}")
        
        # Parse markdown content and convert to docx
        lines = content.split('\n')
        in_code_block = False
        
        for line in lines:
            # Handle headings (## Heading)
            if line.startswith('#'):
                level = len(line.split(' ')[0])  # Count the # symbols
                heading_text = line.lstrip('#').strip()
                doc.add_heading(heading_text, level)
                continue
                
            # Handle issue blocks
            if line.startswith('### Issue'):
                doc.add_heading(line.strip('# '), 3)
                continue
                
            # Handle bold text
            if '**Original**:' in line:
                p = doc.add_paragraph()
                p.add_run('Original: ').bold = True
                p.add_run(line.split('**Original**:')[1].strip())
                continue
                
            if '**Issue**:' in line:
                p = doc.add_paragraph()
                p.add_run('Issue: ').bold = True
                p.add_run(line.split('**Issue**:')[1].strip())
                continue
                
            if '**Suggestion**:' in line:
                p = doc.add_paragraph()
                p.add_run('Suggestion: ').bold = True
                p.add_run(line.split('**Suggestion**:')[1].strip())
                continue
            
            # Handle code blocks
            if line.startswith('```'):
                in_code_block = not in_code_block
                continue
                
            # Handle separators
            if line.startswith('---'):
                doc.add_paragraph('').add_run('─' * 50)
                continue
                
            # Handle regular text
            if line.strip() and not in_code_block:
                doc.add_paragraph(line)
        
        # Save the document
        doc.save(output_path)
        return output_path
    
    def combine_results(self, results, output_path, docx_output_path, total_input_tokens=0, total_output_tokens=0):
        """Combine all segment results into a single document"""
        # Create a combined markdown document
        with open(output_path, "w") as f:
            f.write("# Complete Document Proofreading Report\n\n")
            f.write("## Overview\n\n")
            f.write(f"Total segments processed: {len(results)}\n\n")
            
            # Calculate statistics
            total_issues = 0
            issue_types = {}
            
            for i, result in enumerate(results):
                feedback = result["feedback"]
                # Try to count issues using regex
                issue_matches = re.findall(r'### Issue \d+:', feedback)
                segment_issues = len(issue_matches)
                total_issues += segment_issues
                
                # Try to extract issue types
                type_matches = re.findall(r'### Issue \d+: \[(.*?)\]', feedback)
                for issue_type in type_matches:
                    issue_types[issue_type] = issue_types.get(issue_type, 0) + 1
            
            # Add cost calculation
            costs = calculate_costs(total_input_tokens, total_output_tokens)
            
            # Write cost information
            f.write("### Cost Analysis\n")
            f.write(f"- Total Input Tokens: {costs['input_tokens']:,}\n")
            f.write(f"- Total Output Tokens: {costs['output_tokens']:,}\n")
            f.write(f"- Estimated API Cost: ${costs['total_cost']:.4f}\n\n")
            
            # Write statistics
            f.write(f"Total issues identified: {total_issues}\n\n")
            
            if issue_types:
                f.write("### Issue Types\n")
                for issue_type, count in sorted(issue_types.items(), key=lambda x: x[1], reverse=True):
                    f.write(f"- {issue_type}: {count}\n")
                f.write("\n")
            
            # Write each segment's feedback
            for i, result in enumerate(results):
                f.write(f"## Segment {i+1}\n\n")
                f.write("### Sections Included\n")
                for title in result["section_titles"]:
                    f.write(f"- {title}\n")
                f.write("\n### Feedback\n\n")
                f.write(result["feedback"])
                f.write("\n\n---\n\n")
        
        # Create DOCX version (simplified)
        self.create_docx_report(
            f"# Complete Document Proofreading Report\n\nTotal segments: {len(results)}\nTotal issues: {total_issues}\n\n" +
            "\n\n".join([result["feedback"] for result in results]),
            docx_output_path,
            include_costs=True,
            input_tokens=total_input_tokens,
            output_tokens=total_output_tokens
        )
        
        # Also save JSON format for programmatic use
        combined_json = {
            "segments": results,
            "statistics": {
                "total_segments": len(results),
                "total_issues": total_issues,
                "issue_types": issue_types,
                "tokens": {
                    "input": total_input_tokens,
                    "output": total_output_tokens
                },
                "costs": costs
            },
            "docx_path": docx_output_path
        }
        
        json_path = output_path.replace(".md", ".json")
        with open(json_path, "w") as f:
            json.dump(combined_json, f, indent=2)
        
        return combined_json