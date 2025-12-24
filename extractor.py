import os
import glob
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor

# XML namespaces used in PPTX files
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

# EMU to pixels conversion (96 DPI standard)
EMU_PER_PIXEL = 9525

def emu_to_px(emu):
    """Convert EMU to pixels at 96 DPI."""
    if emu is None:
        return None
    return round(emu / EMU_PER_PIXEL, 1)

def rgb_to_hex(rgb_color):
    """Convert RGBColor to hex string."""
    if rgb_color is None:
        return None
    try:
        return f"#{rgb_color}"
    except:
        return None

def extract_shape_layout(shape):
    """Extract position, size, and rotation from a shape."""
    try:
        return {
            'x': emu_to_px(shape.left),
            'y': emu_to_px(shape.top),
            'width': emu_to_px(shape.width),
            'height': emu_to_px(shape.height),
            'rotation': shape.rotation if shape.rotation else 0
        }
    except Exception as e:
        return None

def extract_fill_style(shape):
    """Extract fill color from a shape."""
    try:
        if not hasattr(shape, 'fill'):
            return None
        fill = shape.fill
        if fill.type is None:
            return None
        
        # Try to get solid fill color
        try:
            fore_color = fill.fore_color
            if fore_color.type is not None:
                # Try RGB first
                try:
                    rgb = fore_color.rgb
                    if rgb:
                        return rgb_to_hex(rgb)
                except:
                    pass
                # Try theme color
                try:
                    theme = fore_color.theme_color
                    brightness = fore_color.brightness
                    return {'theme': str(theme), 'brightness': brightness}
                except:
                    pass
        except:
            pass
    except:
        pass
    return None

def extract_line_style(shape):
    """Extract line/stroke properties from a shape."""
    try:
        if not hasattr(shape, 'line'):
            return None
        line = shape.line
        
        line_data = {}
        
        # Get line width
        if line.width:
            line_data['width'] = emu_to_px(line.width)
        
        # Get line color
        try:
            color = line.color
            if color.type is not None:
                try:
                    rgb = color.rgb
                    if rgb:
                        line_data['color'] = rgb_to_hex(rgb)
                except:
                    try:
                        theme = color.theme_color
                        line_data['theme_color'] = str(theme)
                    except:
                        pass
        except:
            pass
        
        if line_data:
            return line_data
    except:
        pass
    return None

def extract_font_style(shape):
    """Extract font properties from the first text run in a shape."""
    try:
        if not shape.has_text_frame:
            return None
        
        tf = shape.text_frame
        if not tf.paragraphs:
            return None
        
        font_data = {}
        
        # Get alignment from first paragraph
        para = tf.paragraphs[0]
        if para.alignment:
            font_data['alignment'] = str(para.alignment).replace('TEXT_ALIGN.', '').lower()
        
        # Get font properties from first run with text
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text.strip():
                    font = run.font
                    if font.name:
                        font_data['font_name'] = font.name
                    if font.size:
                        # Convert points to CSS pixels (96 DPI / 72 DPI = 1.333)
                        font_data['font_size'] = round(font.size.pt * (96 / 72), 1)
                    if font.bold:
                        font_data['bold'] = font.bold
                    if font.italic:
                        font_data['italic'] = font.italic
                    
                    # Get font color
                    try:
                        fc = font.color
                        if fc.type is not None:
                            try:
                                rgb = fc.rgb
                                if rgb:
                                    font_data['color'] = rgb_to_hex(rgb)
                            except:
                                try:
                                    theme = fc.theme_color
                                    font_data['theme_color'] = str(theme)
                                except:
                                    pass
                    except:
                        pass
                    
                    # Return after first meaningful run
                    if font_data:
                        return font_data
        
        return font_data if font_data else None
    except:
        return None

def get_shape_type_name(shape):
    """Get a simplified shape type name."""
    try:
        shape_type = shape.shape_type
        if shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            # Try to get the auto shape type
            try:
                auto_type = shape.auto_shape_type
                auto_str = str(auto_type).replace('MSO_AUTO_SHAPE_TYPE.', '').lower()
                if 'arrow' in auto_str:
                    return 'arrow'
                return auto_str
            except:
                return 'auto_shape'
        elif shape_type == MSO_SHAPE_TYPE.LINE:
            return 'connector'
        elif shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            return 'text_box'
        elif shape_type == MSO_SHAPE_TYPE.PICTURE:
            return 'picture'
        elif shape_type == MSO_SHAPE_TYPE.GROUP:
            return 'group'
        elif shape_type == MSO_SHAPE_TYPE.FREEFORM:
            return 'freeform'
        else:
            return str(shape_type).replace('MSO_SHAPE_TYPE.', '').lower()
    except:
        return 'unknown'

def extract_connector_path(shape):
    """Extract start and end points for connector shapes."""
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.LINE:
            # For lines, calculate start and end from position and size
            x = emu_to_px(shape.left)
            y = emu_to_px(shape.top)
            w = emu_to_px(shape.width)
            h = emu_to_px(shape.height)
            
            # Determine direction based on the shape's flip properties
            return {
                'start': {'x': x, 'y': y},
                'end': {'x': x + w, 'y': y + h}
            }
    except:
        pass
    return None

def extract_shape_visual_data(shape, z_index):
    """Extract all visual data for a shape."""
    visual_data = {
        'z_index': z_index,
        'shape_type': get_shape_type_name(shape)
    }
    
    # Layout (position, size, rotation)
    layout = extract_shape_layout(shape)
    if layout:
        visual_data['layout'] = layout
    
    # Fill color
    fill = extract_fill_style(shape)
    if fill:
        visual_data['fill'] = fill
    
    # Line/stroke
    line = extract_line_style(shape)
    if line:
        visual_data['line'] = line
    
    # Font properties
    font = extract_font_style(shape)
    if font:
        visual_data['font'] = font
    
    # Connector path (for lines/arrows)
    if visual_data['shape_type'] in ('connector', 'line'):
        path = extract_connector_path(shape)
        if path:
            visual_data['path'] = path
    
    return visual_data

def get_text_from_shape_xml(shape_elem):
    """Extract all text from a shape XML element."""
    texts = []
    for t in shape_elem.findall('.//a:t', NAMESPACES):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()

def get_hyperlink_from_shape_xml(shape_elem):
    """Extract hyperlink action from shape XML element."""
    # Check for hlinkClick on the shape itself (cNvPr)
    cNvPr = shape_elem.find('.//p:cNvPr', NAMESPACES)
    if cNvPr is not None:
        hlinkClick = cNvPr.find('a:hlinkClick', NAMESPACES)
        if hlinkClick is not None:
            action = hlinkClick.get('action', '')
            if 'customshow' in action.lower():
                # Extract custom show ID
                import re
                match = re.search(r'id=(\d+)', action)
                if match:
                    return {'type': 'customshow', 'id': int(match.group(1))}
            elif action:
                return {'type': 'action', 'action': action}
    
    # Check for hyperlinks in text runs
    for hlinkClick in shape_elem.findall('.//a:hlinkClick', NAMESPACES):
        action = hlinkClick.get('action', '')
        if 'customshow' in action.lower():
            import re
            match = re.search(r'id=(\d+)', action)
            if match:
                return {'type': 'customshow', 'id': int(match.group(1))}
        elif action:
            return {'type': 'action', 'action': action}
    
    return None

def parse_animation_sequence(slide_xml_content):
    """Parse animation sequence from slide XML and return ordered list of shape IDs."""
    root = ET.fromstring(slide_xml_content)
    animation_order = []
    
    # Find all spTgt (shape targets) in the timing section
    for spTgt in root.findall('.//p:spTgt', NAMESPACES):
        spid = spTgt.get('spid')
        if spid and spid not in animation_order:
            animation_order.append(spid)
    
    return animation_order

def parse_shapes_from_slide(slide_xml_content):
    """Parse all shapes from slide XML and return dict keyed by shape ID."""
    root = ET.fromstring(slide_xml_content)
    shapes = {}
    
    # Find all sp (shape) elements
    for sp in root.findall('.//p:sp', NAMESPACES):
        nvSpPr = sp.find('p:nvSpPr', NAMESPACES)
        if nvSpPr is not None:
            cNvPr = nvSpPr.find('p:cNvPr', NAMESPACES)
            if cNvPr is not None:
                shape_id = cNvPr.get('id')
                shape_name = cNvPr.get('name', '')
                text = get_text_from_shape_xml(sp)
                hyperlink = get_hyperlink_from_shape_xml(sp)
                
                if shape_id:
                    shapes[shape_id] = {
                        'id': shape_id,
                        'name': shape_name,
                        'text': text,
                        'hyperlink': hyperlink
                    }
    
    # Also find connector/line shapes (cxnSp elements)
    for cxn in root.findall('.//p:cxnSp', NAMESPACES):
        nvCxnSpPr = cxn.find('p:nvCxnSpPr', NAMESPACES)
        if nvCxnSpPr is not None:
            cNvPr = nvCxnSpPr.find('p:cNvPr', NAMESPACES)
            if cNvPr is not None:
                shape_id = cNvPr.get('id')
                shape_name = cNvPr.get('name', '')
                
                if shape_id:
                    shapes[shape_id] = {
                        'id': shape_id,
                        'name': shape_name,
                        'text': '',
                        'hyperlink': None,
                        'is_connector': True
                    }
    
    return shapes

def parse_custom_shows(pptx_path):
    """Parse custom shows from presentation.xml."""
    custom_shows = {}
    
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        try:
            pres_xml = zf.read('ppt/presentation.xml').decode('utf-8')
            root = ET.fromstring(pres_xml)
            
            # Get slide ID to rId mapping
            slide_map = {}
            for sldId in root.findall('.//p:sldId', NAMESPACES):
                slide_id = sldId.get('id')
                r_id = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if slide_id and r_id:
                    slide_map[r_id] = slide_id
            
            # Parse relationships to map rId to slide file
            rels_xml = zf.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
            rels_root = ET.fromstring(rels_xml)
            rid_to_slide = {}
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                r_id = rel.get('Id')
                target = rel.get('Target')
                if target and 'slide' in target.lower() and not 'layout' in target.lower() and not 'master' in target.lower():
                    rid_to_slide[r_id] = target
            
            # Parse custom shows
            for custShow in root.findall('.//p:custShow', NAMESPACES):
                show_name = custShow.get('name', '')
                show_id = custShow.get('id')
                
                if show_id:
                    slides_content = []
                    for sld in custShow.findall('.//p:sld', NAMESPACES):
                        r_id = sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if r_id and r_id in rid_to_slide:
                            slide_file = rid_to_slide[r_id]
                            # Read the slide content
                            try:
                                slide_path = f'ppt/{slide_file}' if not slide_file.startswith('slides/') else f'ppt/{slide_file}'
                                if slide_file.startswith('slides/'):
                                    slide_path = f'ppt/{slide_file}'
                                slide_xml = zf.read(slide_path).decode('utf-8')
                                shapes = parse_shapes_from_slide(slide_xml)
                                # Get all text from shapes with content
                                slide_texts = []
                                for shape in shapes.values():
                                    if shape['text']:
                                        slide_texts.append(shape['text'])
                                slides_content.append({
                                    'slide_file': slide_file,
                                    'texts': slide_texts
                                })
                            except Exception as e:
                                slides_content.append({
                                    'slide_file': slide_file,
                                    'error': str(e)
                                })
                    
                    custom_shows[int(show_id)] = {
                        'name': show_name,
                        'id': int(show_id),
                        'slides': slides_content
                    }
        except Exception as e:
            print(f"Error parsing custom shows: {e}")
    
    return custom_shows

def save_presentation_structure(prs, file_path):
    """Save a simplified representation focusing on animation order and hyperlinks."""
    
    custom_shows = parse_custom_shows(file_path)
    
    # Calculate slide dimensions in pixels (from EMU)
    slide_width = emu_to_px(prs.slide_width)
    slide_height = emu_to_px(prs.slide_height)
    
    presentation_data = {
        "file_path": str(file_path),
        "file_name": Path(file_path).name,
        "total_slides": len(prs.slides),
        "slide_width": slide_width,
        "slide_height": slide_height,
        "custom_shows": custom_shows,
        "slides": []
    }
    
    # Build a mapping of shape ID to pptx shape object for visual data
    with zipfile.ZipFile(file_path, 'r') as zf:
        for slide_num, slide in enumerate(prs.slides, 1):
            slide_file = f'ppt/slides/slide{slide_num}.xml'
            
            # Build shape ID to pptx shape mapping for this slide
            pptx_shapes_by_id = {}
            for z_idx, shape in enumerate(slide.shapes):
                pptx_shapes_by_id[str(shape.shape_id)] = (shape, z_idx)
            
            try:
                slide_xml = zf.read(slide_file).decode('utf-8')
                
                # Get animation order
                animation_order = parse_animation_sequence(slide_xml)
                
                # Get all shapes
                shapes = parse_shapes_from_slide(slide_xml)
                
                # Build ordered animation list
                animation_sequence = []
                sequence_num = 1
                
                for shape_id in animation_order:
                    if shape_id in shapes:
                        shape = shapes[shape_id]
                        # Include all animated shapes (text or not - could be rectangles, decorative shapes)
                        entry = {
                            'sequence': sequence_num,
                            'text': shape['text'] if shape['text'] else '',
                            'shape_name': shape['name']
                        }
                        
                        # Add visual data if available
                        if shape_id in pptx_shapes_by_id:
                            pptx_shape, z_idx = pptx_shapes_by_id[shape_id]
                            visual = extract_shape_visual_data(pptx_shape, z_idx)
                            if visual:
                                for key, value in visual.items():
                                    entry[key] = value
                        
                        # Add hyperlink info if present
                        if shape['hyperlink']:
                            entry['hyperlink'] = shape['hyperlink']
                            # If it's a custom show, include the linked content
                            if shape['hyperlink']['type'] == 'customshow':
                                cs_id = shape['hyperlink']['id']
                                if cs_id in custom_shows:
                                    entry['linked_content'] = custom_shows[cs_id]
                        
                        animation_sequence.append(entry)
                        sequence_num += 1
                
                # Also get shapes that might not be animated (static content)
                static_shapes = []
                animated_ids = set(animation_order)
                for shape_id, shape in shapes.items():
                    if shape_id not in animated_ids:
                        # Include all shapes - text, connectors, or decorative rectangles
                        # Skip if it's truly empty (no text, no visual importance)
                        static_entry = {
                            'text': shape['text'] if shape['text'] else '',
                            'shape_name': shape['name'],
                            'static': True
                        }
                        
                        # Add visual data if available
                        if shape_id in pptx_shapes_by_id:
                            pptx_shape, z_idx = pptx_shapes_by_id[shape_id]
                            visual = extract_shape_visual_data(pptx_shape, z_idx)
                            if visual:
                                for key, value in visual.items():
                                    static_entry[key] = value
                        
                        if shape['hyperlink']:
                            static_entry['hyperlink'] = shape['hyperlink']
                            if shape['hyperlink']['type'] == 'customshow':
                                cs_id = shape['hyperlink']['id']
                                if cs_id in custom_shows:
                                    static_entry['linked_content'] = custom_shows[cs_id]
                        static_shapes.append(static_entry)
                
                slide_info = {
                    'slide_number': slide_num,
                    'animation_sequence': animation_sequence,
                }
                
                if static_shapes:
                    slide_info['static_content'] = static_shapes
                
                presentation_data['slides'].append(slide_info)
                
            except Exception as e:
                presentation_data['slides'].append({
                    'slide_number': slide_num,
                    'error': str(e)
                })
    
    # Save to JSON file
    output_path = Path(file_path).with_suffix('.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(presentation_data, f, indent=2, ensure_ascii=False)
    
    print(f"Presentation structure saved to: {output_path}")
    return output_path


def get_pptx_file():
    script_directory = os.path.dirname(os.path.abspath(__file__))
   
    #expect hsu-pptx or pptx folder to be in the same directory as this script
    path = Path(script_directory).parent / 'hsu-pptx'
    if not path.is_dir():
        path = Path(script_directory).parent / 'pptx'    
    if not path.is_dir():
        print(f"Error. Files not found in: {Path(script_directory).parent}\n" 
              f"Add a folder named 'pptx' to the same directory as this script and add .pptx files to it.")
        exit()
    # Use glob to filter and sort .pptx files
    file_list = sorted(glob.glob(os.path.join(path, '*.pptx')))

    # Print the list of files to the console
    if file_list:
        print(f"Extract from: ${path}")
        for index, file in enumerate(file_list):
            print(f"{index + 1}. {Path(file).name}")

   # Ask the user to select a file
    while True:
        try:
            selection = int(input("Enter the number of the file you want to extract text from (0 to exit): "))
            
            # Check if the selection is valid
            if 0 <= selection <= len(file_list):
                if selection == 0:
                    print("Exiting...")
                    exit()
                else:
                    selected_file = file_list[selection - 1]
                    print(f"Selected: {selected_file}")
                    return selected_file

            else:
                print("Invalid selection. Please enter a valid number.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")

def main():
    import sys
    
    # Accept file path as command-line argument, or fall back to interactive selection
    if len(sys.argv) > 1:
        file_name = sys.argv[1]
        if not os.path.exists(file_name):
            print(f"Error: File not found: {file_name}")
            exit(1)
        if not file_name.endswith('.pptx'):
            print(f"Error: File must be a .pptx file: {file_name}")
            exit(1)
        print(f"Processing: {file_name}")
    else:
        file_name = get_pptx_file()
    
    # Load the presentation
    prs = Presentation(file_name)
    
    save_presentation_structure(prs, file_name)

if __name__ == "__main__":
    main()