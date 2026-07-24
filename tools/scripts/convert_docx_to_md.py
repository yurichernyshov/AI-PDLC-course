#!/usr/bin/env python3
"""Convert a .docx file to Markdown, extracting images with readable names."""

import os
import sys
import zipfile
import shutil
import xml.etree.ElementTree as ET
import re
import hashlib


NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}


def slugify(text, max_len=40):
    """Create a filesystem-safe name from text."""
    text = text.strip().lower()
    text = re.sub(r'[^\w\s-]', '', text)
    text = re.sub(r'[\s]+', '-', text)
    text = re.sub(r'[-]+', '-', text).strip('-')
    if not text:
        return 'image'
    return text[:max_len]


def extract_text_with_style(paragraph_elem):
    """Extract text from a paragraph element with markdown formatting."""
    texts = []
    for run in paragraph_elem.findall('.//w:r', NS):
        # Skip images in text extraction
        if run.find('.//a:blip', NS) is not None or run.find('.//w:drawing', NS) is not None:
            continue
        t_elem = run.find('.//w:t', NS)
        if t_elem is not None and t_elem.text:
            text = t_elem.text
            
            # Check for bold/italic
            rPr = run.find('w:rPr', NS)
            is_bold = False
            is_italic = False
            if rPr is not None:
                bold_elem = rPr.find('w:b', NS)
                italic_elem = rPr.find('w:i', NS)
                is_bold = bold_elem is not None and bold_elem.get(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true') != 'false'
                is_italic = italic_elem is not None and italic_elem.get(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true') != 'false'
            
            if is_bold and is_italic:
                text = f'***{text}***'
            elif is_bold:
                text = f'**{text}**'
            elif is_italic:
                text = f'*{text}*'
            
            texts.append(text)
    
    return ''.join(texts)


def get_heading_level(paragraph_elem):
    """Get heading level from paragraph style."""
    pPr = paragraph_elem.find('w:pPr', NS)
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is not None:
            val = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            if val.startswith('Heading'):
                try:
                    return int(val.replace('Heading', ''))
                except ValueError:
                    pass
    return None


def is_list_item(paragraph_elem, numbering_map):
    """Check if paragraph is a list item."""
    pPr = paragraph_elem.find('w:pPr', NS)
    if pPr is not None:
        numPr = pPr.find('w:numPr', NS)
        if numPr is not None:
            numId_elem = numPr.find('w:numId', NS)
            ilvl_elem = numPr.find('w:ilvl', NS)
            if numId_elem is not None:
                numId = numId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                ilvl = ilvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0') if ilvl_elem is not None else '0'
                # Determine if numbered or bulleted
                num_info = numbering_map.get(numId, {})
                if num_info.get('type') == 'bullet':
                    return ('bullet', int(ilvl))
                else:
                    return ('numbered', int(ilvl), num_info)
    return None


def extract_images_from_paragraph(paragraph_elem, doc_path, images_dir, image_counter, doc_rels):
    """Extract images from a paragraph and return markdown image refs."""
    refs = []
    for drawing in paragraph_elem.findall('.//w:drawing', NS):
        blip = drawing.find('.//a:blip', NS)
        if blip is None:
            continue
        
        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if embed_id is None:
            continue
        
        # Find the image file from relationships
        doc_dir = os.path.dirname(doc_path)
        rel_elem = doc_rels.find(f'.//*[@Id="{embed_id}"]')
        if rel_elem is None:
            continue
        
        target = rel_elem.get('Target', '')
        # Resolve relative path
        image_path = os.path.normpath(os.path.join(doc_dir, target))
        
        if not os.path.exists(image_path):
            continue
        
        # Get extension
        ext = os.path.splitext(image_path)[1].lower()
        if ext not in ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'):
            ext = '.png'
        
        # Try to get a description from the blip
        desc = blip.get('{http://schemas.openxmlformats.org/drawingml/2006/main}descr', '')
        if not desc:
            # Try to find alt text from pic element
            pic_elem = drawing.find('.//pic:nvPicPr/pic:cNvPr', NS)
            if pic_elem is not None:
                desc = pic_elem.get('descr', '') or pic_elem.get('title', '')
        
        if not desc:
            # Use file name without extension
            desc = os.path.splitext(os.path.basename(image_path))[0]
        
        # Create unique readable name
        image_counter[0] += 1
        slug = slugify(desc)
        filename = f"image-{image_counter[0]:02d}-{slug}{ext}"
        
        dest_path = os.path.join(images_dir, filename)
        shutil.copy2(image_path, dest_path)
        
        refs.append(f"![{desc}](images/{filename})")
    
    return refs


def convert_docx_to_md(docx_path, output_md_path, images_dir):
    """Main conversion function."""
    os.makedirs(images_dir, exist_ok=True)
    
    image_counter = [0]  # mutable counter
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        # Load document.xml
        doc_xml = z.read('word/document.xml')
        root = ET.fromstring(doc_xml)
        
        # Load relationships for images
        try:
            rels_xml = z.read('word/_rels/document.xml.rels')
            doc_rels = ET.fromstring(rels_xml)
        except Exception:
            doc_rels = ET.Element('Relationships')
        
        # Load numbering definitions
        try:
            numbering_xml = z.read('word/numbering.xml')
            numbering_root = ET.fromstring(numbering_xml)
            numbering_map = {}
            for num in numbering_root.findall('.//w:num', NS):
                numId = num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
                abstractNumId = num.find('.//w:abstractNumId', NS)
                if abstractNumId is not None:
                    absId = abstractNumId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    # Find the abstract numbering definition
                    for abstract in numbering_root.findall('.//w:abstractNum', NS):
                        if abstract.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId') == absId:
                            lvl = abstract.find('.//w:lvl', NS)
                            if lvl is not None:
                                numFmt = lvl.find('w:numFmt', NS)
                                if numFmt is not None:
                                    fmt = numFmt.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
                                    numbering_map[numId] = {'type': 'numbered' if fmt else 'bullet'}
                                else:
                                    numbering_map[numId] = {'type': 'bullet'}
                            break
            print(f"Numbering map: {numbering_map}")
        except Exception as e:
            print(f"Could not load numbering: {e}")
            numbering_map = {}
    
    # Now parse the document body
    body = root.find('w:body', NS)
    if body is None:
        print("No body found in document")
        return
    
    lines = []
    list_counters = {}  # track numbered list counters per level/numId
    
    paragraphs = body.findall('w:p', NS)
    
    for para in paragraphs:
        # Check for heading
        heading_level = get_heading_level(para)
        
        # Check for list item
        list_info = is_list_item(para, numbering_map)
        
        text = extract_text_with_style(para)
        
        # Extract images
        image_refs = extract_images_from_paragraph(para, docx_path, images_dir, image_counter, doc_rels)
        
        # Add image references
        if image_refs:
            if text:
                lines.append(text)
            lines.extend(image_refs)
            continue
        
        if heading_level is not None:
            prefix = '#' * heading_level
            lines.append(f"\n{prefix} {text}")
        elif list_info:
            list_type = list_info[0]
            level = list_info[1]
            indent = '  ' * level
            
            if list_type == 'bullet':
                lines.append(f"{indent}- {text}")
            else:
                # Numbered list
                key = f"num-{level}"
                list_counters[key] = list_counters.get(key, 0) + 1
                lines.append(f"{indent}{list_counters[key]}. {text}")
        else:
            # Check if it's a table row (skip, handle separately)
            if para.find('.//w:tbl', NS) is not None:
                continue
            # Regular paragraph
            if text.strip():
                lines.append(text)
            else:
                # Empty line
                lines.append('')
    
    # Process tables
    tables = body.findall('.//w:tbl', NS)
    for table in tables:
        lines.append('')
        rows = table.findall('w:tr', NS)
        for i, row in enumerate(rows):
            cells = row.findall('w:tc', NS)
            cell_texts = []
            for cell in cells:
                cell_text = ' '.join(
                    t.text for t in cell.findall('.//w:t', NS) if t.text
                ).strip()
                cell_texts.append(cell_text)
            lines.append('| ' + ' | '.join(cell_texts) + ' |')
            if i == 0:
                lines.append('| ' + ' | '.join(['---'] * len(cell_texts)) + ' |')
        lines.append('')
    
    # Clean up multiple blank lines
    result = '\n'.join(lines)
    result = re.sub(r'\n{3,}', '\n\n', result)
    
    with open(output_md_path, 'w', encoding='utf-8') as f:
        f.write(result)
    
    print(f"Converted: {docx_path}")
    print(f"Output: {output_md_path}")
    print(f"Images saved to: {images_dir}")
    print(f"Total images: {image_counter[0]}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: convert_docx_to_md.py <file.docx>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    base_dir = os.path.dirname(docx_path) or '.'
    filename = os.path.basename(docx_path)
    name_no_ext = os.path.splitext(filename)[0]
    
    output_md = os.path.join(base_dir, f"{name_no_ext}.md")
    images_dir = os.path.join(base_dir, 'images')
    
    convert_docx_to_md(docx_path, output_md, images_dir)
