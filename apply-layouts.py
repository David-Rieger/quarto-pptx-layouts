import win32com.client
import os
import re
import sys

def process_presentation(ppt_path):
    abs_path = os.path.abspath(ppt_path)
    if not os.path.exists(abs_path):
        print(f"Error: File not found '{abs_path}'")
        sys.exit(1)
        
    print(f"Processing PowerPoint: {abs_path}")
    
    app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = app.Presentations.Open(abs_path, WithWindow=False)
    
    try:
        slide_master = presentation.SlideMaster
        
        custom_layouts = {}
        for layout in slide_master.CustomLayouts:
            custom_layouts[layout.Name] = layout
            
        layout_regex = re.compile(r'\[layout:\s*([^\]]+)\]')
        size_regex = re.compile(r'\[size:\s*([^\]]+)\]')
        
        modified = False
        
        for i, slide in enumerate(presentation.Slides):
            notes_page = slide.NotesPage
            
            target_layout_name = None
            target_shape_layout = None
            original_notes_text_layout = None
            
            target_size_name = None
            target_shape_size = None
            original_notes_text_size = None
            
            for shape in notes_page.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    
                    match_layout = layout_regex.search(text)
                    if match_layout:
                        target_layout_name = match_layout.group(1).strip()
                        target_shape_layout = shape
                        original_notes_text_layout = text
                        
                    match_size = size_regex.search(text)
                    if match_size:
                        target_size_name = match_size.group(1).strip()
                        target_shape_size = shape
                        original_notes_text_size = text
                        
            if target_size_name:
                current_layout_name = slide.CustomLayout.Name
                base_name = current_layout_name.replace(" smaller", "").replace(" smallest", "").strip()
                desired_layout_name = f"{base_name} {target_size_name}"
                
                print(f"Slide {i+1}: Size '{target_size_name}' requested. Changing from '{current_layout_name}' to '{desired_layout_name}'")
                
                if desired_layout_name in custom_layouts:
                    slide.CustomLayout = custom_layouts[desired_layout_name]
                    clean_text = size_regex.sub('', original_notes_text_size).strip()
                    target_shape_size.TextFrame.TextRange.Text = clean_text
                    modified = True
                else:
                    print(f"Slide {i+1}: Warning - Layout '{desired_layout_name}' not found!")
            
            elif target_layout_name:
                print(f"Slide {i+1}: Explicit layout '{target_layout_name}' requested.")
                if target_layout_name in custom_layouts:
                    slide.CustomLayout = custom_layouts[target_layout_name]
                    clean_text = layout_regex.sub('', original_notes_text_layout).strip()
                    target_shape_layout.TextFrame.TextRange.Text = clean_text
                    modified = True
                else:
                    print(f"Slide {i+1}: Warning - Layout '{target_layout_name}' not found!")
        
        if modified:
            presentation.Save()
            print("Presentation successfully updated.")
        else:
            print("No layout changes needed.")
            
    finally:
        presentation.Close()
        app.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python apply-layouts.py <path_to_pptx>")
        sys.exit(1)
        
    process_presentation(sys.argv[1])
