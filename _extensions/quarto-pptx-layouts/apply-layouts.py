import os
import sys
import zipfile
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET

# Namespaces required for parsing PowerPoint XML
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
}
ET.register_namespace('p', NAMESPACES['p'])
ET.register_namespace('a', NAMESPACES['a'])
ET.register_namespace('r', NAMESPACES['r'])
ET.register_namespace('', NAMESPACES['rel'])

def process_presentation(ppt_path):
    abs_path = os.path.abspath(ppt_path)
    if not os.path.exists(abs_path):
        print(f"Error: File not found '{abs_path}'")
        return
        
    print(f"Processing PowerPoint: {abs_path}")
    
    # We work in a temporary directory to unpack the PPTX, modify it, and repack it
    temp_dir = tempfile.mkdtemp()
    try:
        # Extract pptx
        with zipfile.ZipFile(abs_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        # 1. Map layouts
        layouts_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        if not os.path.exists(layouts_dir):
            print("No slide layouts found.")
            return
            
        layout_map = {} # name -> filename (e.g. 'Title and Content' -> 'slideLayout1.xml')
        layout_name_to_target = {} # name -> '../slideLayouts/slideLayout1.xml'
        
        for filename in os.listdir(layouts_dir):
            if filename.endswith('.xml'):
                tree = ET.parse(os.path.join(layouts_dir, filename))
                root = tree.getroot()
                # Find <p:cSld name="...">
                cSld = root.find('.//p:cSld', NAMESPACES)
                if cSld is not None and 'name' in cSld.attrib:
                    name = cSld.attrib['name']
                    layout_map[name] = filename
                    layout_name_to_target[name] = f"../slideLayouts/{filename}"
                    
        # 2. Process slides
        slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
        if not os.path.exists(slides_dir):
            print("No slides found.")
            return
            
        layout_regex = re.compile(r'\[layout:\s*([^\]]+)\]')
        size_regex = re.compile(r'\[size:\s*([^\]]+)\]')
        
        modified = False
        
        for filename in os.listdir(slides_dir):
            if not filename.endswith('.xml'):
                continue
                
            slide_idx = re.search(r'\d+', filename)
            if not slide_idx: continue
            idx = slide_idx.group(0)
            
            # Read rels
            rels_file = os.path.join(slides_dir, '_rels', f"{filename}.rels")
            if not os.path.exists(rels_file):
                continue
                
            tree_rels = ET.parse(rels_file)
            root_rels = tree_rels.getroot()
            
            # Find notes slide relationship and current layout
            notes_target = None
            layout_rel = None
            current_layout_target = None
            
            for rel in root_rels.findall('.//rel:Relationship', NAMESPACES):
                rel_type = rel.attrib.get('Type', '')
                if rel_type.endswith('/notesSlide'):
                    notes_target = rel.attrib.get('Target')
                elif rel_type.endswith('/slideLayout'):
                    layout_rel = rel
                    current_layout_target = rel.attrib.get('Target')
                    
            if not notes_target or layout_rel is None:
                continue
                
            # Read notes slide
            notes_path = os.path.normpath(os.path.join(slides_dir, notes_target))
            if not os.path.exists(notes_path):
                continue
                
            tree_notes = ET.parse(notes_path)
            root_notes = tree_notes.getroot()
            
            # Extract text from notes
            notes_text = ""
            for t in root_notes.findall('.//a:t', NAMESPACES):
                if t.text: notes_text += t.text + " "
                
            target_layout_name = None
            target_size_name = None
            
            match_layout = layout_regex.search(notes_text)
            if match_layout:
                target_layout_name = match_layout.group(1).strip()
                
            match_size = size_regex.search(notes_text)
            if match_size:
                target_size_name = match_size.group(1).strip()
                
            if target_size_name:
                # Need to find the current layout name to append the size
                current_layout_name = None
                for name, target in layout_name_to_target.items():
                    if target == current_layout_target:
                        current_layout_name = name
                        break
                        
                if current_layout_name:
                    base_name = current_layout_name.replace(" smallest", "").replace(" smaller", "").strip()
                    desired_layout_name = f"{base_name} {target_size_name}"
                    
                    if desired_layout_name in layout_name_to_target:
                        layout_rel.attrib['Target'] = layout_name_to_target[desired_layout_name]
                        print(f"Slide {idx}: Size '{target_size_name}' requested. Changing from '{current_layout_name}' to '{desired_layout_name}'")
                        tree_rels.write(rels_file, xml_declaration=True, encoding='UTF-8')
                        modified = True
                        
                        # Remove the tag from the notes
                        for t in root_notes.findall('.//a:t', NAMESPACES):
                            if t.text and match_size.group(0) in t.text:
                                t.text = t.text.replace(match_size.group(0), "")
                        tree_notes.write(notes_path, xml_declaration=True, encoding='UTF-8')
                    else:
                        print(f"Slide {idx}: Warning - Layout '{desired_layout_name}' not found!")
                        
            elif target_layout_name:
                if target_layout_name in layout_name_to_target:
                    layout_rel.attrib['Target'] = layout_name_to_target[target_layout_name]
                    print(f"Slide {idx}: Explicit layout '{target_layout_name}' requested.")
                    tree_rels.write(rels_file, xml_declaration=True, encoding='UTF-8')
                    modified = True
                    
                    # Remove the tag from the notes
                    for t in root_notes.findall('.//a:t', NAMESPACES):
                        if t.text and match_layout.group(0) in t.text:
                            t.text = t.text.replace(match_layout.group(0), "")
                    tree_notes.write(notes_path, xml_declaration=True, encoding='UTF-8')
                else:
                    print(f"Slide {idx}: Warning - Layout '{target_layout_name}' not found!")
                    
        if modified:
            # Package back to zip
            temp_zip_path = abs_path + ".tmp.pptx"
            with zipfile.ZipFile(temp_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zip_ref.write(file_path, arcname)
                        
            # Replace original
            shutil.move(temp_zip_path, abs_path)
            print("Presentation successfully updated.")
        else:
            print("No layout changes needed.")
            
    finally:
        shutil.rmtree(temp_dir)

def main():
    # If a direct file is passed via CLI
    if len(sys.argv) > 1:
        process_presentation(sys.argv[1])
        return

    # Check if executed as a Quarto post-render script
    quarto_files = os.environ.get("QUARTO_PROJECT_OUTPUT_FILES", "")
    
    if quarto_files:
        files = [f.strip() for f in quarto_files.split('\n') if f.strip()]
        for f in files:
            if f.lower().endswith('.pptx'):
                process_presentation(f)
        return
        
    print("Usage: python apply-layouts.py <path_to_pptx>")
    print("Or run within a Quarto project using the post-render hook.")

if __name__ == "__main__":
    main()
