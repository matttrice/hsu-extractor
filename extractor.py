import os
import glob
import json
import math
import sys
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

# Target canvas dimensions for MBS (960×540 pixels, 16:9 aspect ratio)
TARGET_CANVAS_WIDTH = 960
TARGET_CANVAS_HEIGHT = 540

# EMU to pixels conversion (96 DPI standard)
EMU_PER_PIXEL = 9525

def emu_to_px(emu):
    """Convert EMU to pixels at 96 DPI."""
    if emu is None:
        return None
    return round(emu / EMU_PER_PIXEL, 1)

def scale_to_target(value, source_width, target_width=TARGET_CANVAS_WIDTH):
    """Scale a coordinate value from source slide dimensions to target canvas dimensions.
    
    Args:
        value: The coordinate value to scale (can be None)
        source_width: The width of the source PowerPoint slide in pixels
        target_width: The target canvas width (default: 960)
    
    Returns:
        Scaled value rounded to 1 decimal place, or None if value is None
    """
    if value is None:
        return None
    if source_width == target_width:
        return round(value, 1)
    scale_factor = target_width / source_width
    return round(value * scale_factor, 1)

def rgb_to_hex(rgb_color):
    """Convert RGBColor to hex string."""
    if rgb_color is None:
        return None
    try:
        return f"#{rgb_color}"
    except:
        return None

def enumerate_shapes_recursive(shapes, z_index_start=0, parent_group_id=None):
    """Recursively enumerate all shapes including those inside groups.
    
    Args:
        shapes: Collection of shapes to enumerate (from slide.shapes or group.shapes)
        z_index_start: Starting z-index for enumeration
        parent_group_id: Parent group's shape ID if applicable
    
    Yields:
        Tuple of (z_index, shape, group_id) for each shape found
    """
    z_idx = z_index_start
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # This is a group - recursively enumerate its children
            group_id = str(shape.shape_id)
            yield from enumerate_shapes_recursive(shape.shapes, z_idx, group_id)
            # Count group members for z-index offset
            z_idx += len(list(shape.shapes))
        else:
            # Regular shape
            yield (z_idx, shape, parent_group_id)
            z_idx += 1

def extract_layout_from_xml(shape_elem, slide_width=None):
    """Extract layout from shape XML element.
    
    Args:
        shape_elem: XML element (p:sp or p:cxnSp)
        slide_width: Source slide width for auto-scaling
    
    Returns:
        Layout dict with coordinates scaled to target canvas
    """
    try:
        spPr = shape_elem.find('p:spPr', NAMESPACES)
        if spPr is None:
            return None
        
        xfrm = spPr.find('a:xfrm', NAMESPACES)
        if xfrm is None:
            return None
        
        # Get offset (position)
        off = xfrm.find('a:off', NAMESPACES)
        # Get extents (size)
        ext = xfrm.find('a:ext', NAMESPACES)
        
        if off is None or ext is None:
            return None
        
        # Extract EMU values and convert to pixels
        x = emu_to_px(int(off.get('x', 0)))
        y = emu_to_px(int(off.get('y', 0)))
        width = emu_to_px(int(ext.get('cx', 0)))
        height = emu_to_px(int(ext.get('cy', 0)))
        rotation = int(xfrm.get('rot', 0)) / 60000  # Convert from 1/60000 degrees
        
        # Scale to target canvas
        if slide_width is not None:
            x = scale_to_target(x, slide_width)
            y = scale_to_target(y, slide_width)
            width = scale_to_target(width, slide_width)
            height = scale_to_target(height, slide_width)
        
        return {
            'x': x,
            'y': y,
            'width': width,
            'height': height,
            'rotation': rotation
        }
    except:
        return None

def extract_font_from_xml(shape_elem):
    """Extract font properties from shape XML element.
    
    Args:
        shape_elem: XML element (p:sp)
    
    Returns:
        Font dict with properties
    """
    try:
        txBody = shape_elem.find('.//p:txBody', NAMESPACES)
        if txBody is None:
            return None
        
        font_data = {}
        
        # Get bodyPr for vertical alignment
        bodyPr = txBody.find('a:bodyPr', NAMESPACES)
        if bodyPr is not None:
            anchor = bodyPr.get('anchor')
            if anchor == 't':
                font_data['v_align'] = 'top'
            elif anchor == 'ctr':
                font_data['v_align'] = 'middle'
            elif anchor == 'b':
                font_data['v_align'] = 'bottom'
        
        # Get first paragraph for alignment
        para = txBody.find('a:p', NAMESPACES)
        if para is not None:
            pPr = para.find('a:pPr', NAMESPACES)
            if pPr is not None:
                algn = pPr.get('algn')
                if algn == 'l':
                    font_data['align'] = 'left'
                elif algn == 'ctr':
                    font_data['align'] = 'center'
                elif algn == 'r':
                    font_data['align'] = 'right'
            
            # Get first run for font properties
            r = para.find('.//a:r', NAMESPACES)
            if r is not None:
                rPr = r.find('a:rPr', NAMESPACES)
                if rPr is not None:
                    # Font size (in 1/100 points, convert to CSS pixels)
                    # Formula: (sz/100) * (96/72) = sz/100 * 1.333...
                    sz = rPr.get('sz')
                    if sz:
                        points = int(sz) / 100
                        font_size = round(points * 1.333, 1)
                        # Apply canvas scale factor (same as coordinate scaling)
                        # This ensures fonts are proportional to the scaled canvas
                        # Note: This requires slide_width context which XML extraction doesn't have
                        # For now, fonts from XML extraction won't be pre-scaled
                        # (They'll be manually scaled if needed in Svelte)
                        font_data['font_size'] = font_size
                    
                    # Bold
                    if rPr.get('b') == '1':
                        font_data['bold'] = True
                    
                    # Italic
                    if rPr.get('i') == '1':
                        font_data['italic'] = True
                    
                    # Font name
                    latin = rPr.find('a:latin', NAMESPACES)
                    if latin is not None:
                        typeface = latin.get('typeface')
                        if typeface:
                            font_data['font_name'] = typeface
                            # Check if font name includes "Bold" and set bold property
                            if 'bold' in typeface.lower():
                                font_data['bold'] = True
                    
                    # Color
                    solidFill = rPr.find('a:solidFill', NAMESPACES)
                    if solidFill is not None:
                        srgbClr = solidFill.find('a:srgbClr', NAMESPACES)
                        if srgbClr is not None:
                            val = srgbClr.get('val')
                            if val:
                                font_data['color'] = f"#{val}"
        
        return font_data if font_data else None
    except:
        return None

def extract_visual_data_from_xml(shape_elem, z_index, slide_width=None):
    """Extract visual data from shape XML element.
    
    Args:
        shape_elem: XML element (p:sp or p:cxnSp)
        z_index: The z-index for this shape
        slide_width: Source slide width for auto-scaling
    
    Returns:
        Visual data dict
    """
    visual_data = {'z_index': z_index}
    
    # Determine shape type from XML
    if shape_elem.tag.endswith('sp'):
        visual_data['shape_type'] = 'text_box'
    elif shape_elem.tag.endswith('cxnSp'):
        visual_data['shape_type'] = 'connector'
    
    # Extract layout
    layout = extract_layout_from_xml(shape_elem, slide_width)
    if layout:
        visual_data['layout'] = layout
    
    # Extract font
    font = extract_font_from_xml(shape_elem)
    if font:
        visual_data['font'] = font
    
    # TODO: Could also extract fill and line from XML if needed
    # For now, these are less critical for grouped text shapes
    
    return visual_data

def extract_shape_layout(shape, slide_width=None):
    """Extract position, size, and rotation from a shape.
    
    Args:
        shape: The pptx shape object
        slide_width: Source slide width in pixels for auto-scaling (optional)
    
    Returns:
        Layout dict with coordinates scaled to target canvas (960×540)
    """
    try:
        # Extract raw coordinates in source dimensions
        x = emu_to_px(shape.left)
        y = emu_to_px(shape.top)
        width = emu_to_px(shape.width)
        height = emu_to_px(shape.height)
        
        # Auto-scale to target canvas if slide_width provided
        if slide_width is not None:
            x = scale_to_target(x, slide_width)
            y = scale_to_target(y, slide_width)  # Use same scale for y
            width = scale_to_target(width, slide_width)
            height = scale_to_target(height, slide_width)
        
        return {
            'x': x,
            'y': y,
            'width': width,
            'height': height,
            'rotation': shape.rotation if shape.rotation else 0
        }
    except Exception as e:
        return None

def calculate_line_endpoints(layout, slide_width=None):
    """Calculate actual line endpoints from layout with rotation.
    
    PowerPoint stores lines as rectangles with rotation. The line runs from
    top-center to bottom-center of the unrotated rectangle, then the whole
    thing is rotated around the center.
    
    For lines:
    - rotation 0: vertical line (top to bottom)
    - rotation 90 or 270: horizontal line (left to right or right to left)
    - other angles: diagonal line
    
    Args:
        layout: Layout dict with x, y, width, height, rotation (already scaled if slide_width was used)
        slide_width: Not used here - layout is already scaled by extract_shape_layout
    
    Returns dict with 'from' and 'to' points {x, y} or None if not applicable.
    """
    if not layout:
        return None
    
    x = layout.get('x', 0)
    y = layout.get('y', 0)
    w = layout.get('width', 0)
    h = layout.get('height', 0)
    rotation = layout.get('rotation', 0)
    
    # Center of the shape
    cx = x + w / 2
    cy = y + h / 2
    
    # Original endpoints (before rotation) - line from top-center to bottom-center
    # The "height" is the line length in its unrotated state
    p1_x, p1_y = cx, y           # top-center
    p2_x, p2_y = cx, y + h       # bottom-center
    
    # Rotate around center
    rad = math.radians(rotation)
    cos_r = math.cos(rad)
    sin_r = math.sin(rad)
    
    def rotate_point(px, py):
        dx = px - cx
        dy = py - cy
        rx = cx + dx * cos_r - dy * sin_r
        ry = cy + dx * sin_r + dy * cos_r
        return round(rx, 1), round(ry, 1)
    
    from_pt = rotate_point(p1_x, p1_y)
    to_pt = rotate_point(p2_x, p2_y)
    
    return {
        'from': {'x': from_pt[0], 'y': from_pt[1]},
        'to': {'x': to_pt[0], 'y': to_pt[1]}
    }

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
        
        # Get dash style from line.dash_style
        try:
            dash_style = line.dash_style
            if dash_style is not None:
                from pptx.enum.dml import MSO_LINE_DASH_STYLE
                dash_map = {
                    MSO_LINE_DASH_STYLE.SOLID: None,  # Don't include for solid
                    MSO_LINE_DASH_STYLE.DASH: 'dash',
                    MSO_LINE_DASH_STYLE.DASH_DOT: 'dashDot',
                    MSO_LINE_DASH_STYLE.DASH_DOT_DOT: 'dashDotDot',
                    MSO_LINE_DASH_STYLE.LONG_DASH: 'lgDash',
                    MSO_LINE_DASH_STYLE.LONG_DASH_DOT: 'lgDashDot',
                    MSO_LINE_DASH_STYLE.ROUND_DOT: 'dot',
                    MSO_LINE_DASH_STYLE.SQUARE_DOT: 'sysDash',  # sysDash maps to SQUARE_DOT in python-pptx
                }
                dash_str = dash_map.get(dash_style)
                if dash_str:
                    line_data['dash'] = dash_str
        except:
            pass
        
        if line_data:
            return line_data
    except:
        pass
    return None


def extract_arrow_ends_from_xml(slide_xml_content, shape_id):
    """Extract arrow head/tail end types from slide XML.
    
    Returns dict with 'headEnd' and 'tailEnd' if present.
    Values are: 'none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'
    """
    try:
        root = ET.fromstring(slide_xml_content)
        
        # Find the shape with matching ID (check both sp and cxnSp elements)
        for sp in root.findall('.//p:sp', NAMESPACES):
            cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
            if cNvPr is not None and cNvPr.get('id') == str(shape_id):
                spPr = sp.find('p:spPr', NAMESPACES)
                if spPr is not None:
                    return _extract_line_ends(spPr)
        
        # Also check connector shapes
        for cxnSp in root.findall('.//p:cxnSp', NAMESPACES):
            cNvPr = cxnSp.find('.//p:cNvPr', NAMESPACES)
            if cNvPr is not None and cNvPr.get('id') == str(shape_id):
                spPr = cxnSp.find('p:spPr', NAMESPACES)
                if spPr is not None:
                    return _extract_line_ends(spPr)
    except:
        pass
    return None


def _extract_line_ends(spPr):
    """Extract headEnd and tailEnd from a spPr element."""
    ln = spPr.find('a:ln', NAMESPACES)
    if ln is None:
        return None
    
    result = {}
    
    headEnd = ln.find('a:headEnd', NAMESPACES)
    if headEnd is not None:
        head_type = headEnd.get('type', 'none')
        if head_type and head_type != 'none':
            result['headEnd'] = head_type
    
    tailEnd = ln.find('a:tailEnd', NAMESPACES)
    if tailEnd is not None:
        tail_type = tailEnd.get('type', 'none')
        if tail_type and tail_type != 'none':
            result['tailEnd'] = tail_type
    
    return result if result else None

def extract_font_style(shape, slide_width=None):
    """Extract font properties from the first text run in a shape.
    
    Args:
        shape: The pptx shape object
        slide_width: Source slide width for auto-scaling font sizes (optional)
    """
    try:
        if not shape.has_text_frame:
            return None
        
        tf = shape.text_frame
        if not tf.paragraphs:
            return None
        
        font_data = {}
        
        # Get vertical alignment from text frame's vertical_anchor property
        try:
            from pptx.enum.text import MSO_ANCHOR
            v_anchor = tf.vertical_anchor
            if v_anchor == MSO_ANCHOR.TOP:
                font_data['v_align'] = 'top'
            elif v_anchor == MSO_ANCHOR.MIDDLE:
                font_data['v_align'] = 'middle'
            elif v_anchor == MSO_ANCHOR.BOTTOM:
                font_data['v_align'] = 'bottom'
        except:
            pass
        
        # Get text wrapping property from text frame
        try:
            # word_wrap is a boolean property on the text_frame
            # True = text wraps, False = text doesn't wrap
            if hasattr(tf, 'word_wrap') and tf.word_wrap:
                font_data['wrap'] = True
        except:
            pass
        
        # Get horizontal alignment from first paragraph
        para = tf.paragraphs[0]
        if para.alignment:
            from pptx.enum.text import PP_ALIGN
            align_map = {
                PP_ALIGN.LEFT: 'left',
                PP_ALIGN.CENTER: 'center',
                PP_ALIGN.RIGHT: 'right',
            }
            if para.alignment in align_map:
                font_data['align'] = align_map[para.alignment]
        
        # Get font properties from first run with text
        found_font_size = False
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text.strip():
                    font = run.font
                    if font.name:
                        font_data['font_name'] = font.name
                        # Check if font name includes "Bold" and set bold property
                        if 'bold' in font.name.lower():
                            font_data['bold'] = True
                    if font.size:
                        # font.size.pt is in PowerPoint points (1/72 inch)
                        # CSS pixels are at 96 DPI, so conversion is: points × (96/72) = points × 1.333...
                        # This is the standard DPI conversion formula and is correct.
                        font_data['font_size'] = round(font.size.pt * (96 / 72), 1)
                        # Apply canvas scale factor (same as coordinate scaling)
                        # This ensures fonts are proportional to the scaled canvas
                        if slide_width is not None:
                            scale_factor = TARGET_CANVAS_WIDTH / slide_width
                            font_data['font_size'] = round(font_data['font_size'] * scale_factor, 1)
                        found_font_size = True
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

def extract_arc_path_from_xml(slide_xml_content, shape_id, shape_layout, slide_width=None):
    """Extract arc path data for a freeform shape from slide XML.
    
    Args:
        slide_xml_content: The slide XML content
        shape_id: The shape ID to extract
        shape_layout: The shape layout (already scaled if slide_width was used)
        slide_width: Not used here - shape_layout is already scaled
    
    Returns arc parameters:
    - from: start point in canvas coordinates (scaled)
    - to: end point in canvas coordinates (scaled)
    - curve: vertical offset for quadratic bezier (negative = up, positive = down)
    - flip: whether the arc is horizontally flipped
    """
    try:
        root = ET.fromstring(slide_xml_content)
        
        # Find the shape with matching ID
        for sp in root.findall('.//p:sp', NAMESPACES):
            cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
            if cNvPr is not None and cNvPr.get('id') == str(shape_id):
                # Check if it's a freeform with custom geometry
                spPr = sp.find('p:spPr', NAMESPACES)
                if spPr is None:
                    continue
                
                custGeom = spPr.find('a:custGeom', NAMESPACES)
                if custGeom is None:
                    continue
                
                # Get transform info
                xfrm = spPr.find('a:xfrm', NAMESPACES)
                flipH = xfrm.get('flipH') == '1' if xfrm is not None else False
                flipV = xfrm.get('flipV') == '1' if xfrm is not None else False
                
                # Find the first path (the stroke path, not fill path)
                pathLst = custGeom.find('a:pathLst', NAMESPACES)
                if pathLst is None:
                    continue
                
                # Get the path that has fill="none" (stroke path)
                stroke_path = None
                for path in pathLst.findall('a:path', NAMESPACES):
                    if path.get('fill') == 'none':
                        stroke_path = path
                        break
                
                if stroke_path is None:
                    # Fall back to first path
                    stroke_path = pathLst.find('a:path', NAMESPACES)
                
                if stroke_path is None:
                    continue
                
                # Get path dimensions for coordinate scaling
                path_w = int(stroke_path.get('w', '21600'))
                path_h = int(stroke_path.get('h', '21600'))
                
                # Extract all points from the path
                points = []
                
                # moveTo is the start point
                moveTo = stroke_path.find('a:moveTo', NAMESPACES)
                if moveTo is not None:
                    pt = moveTo.find('a:pt', NAMESPACES)
                    if pt is not None:
                        points.append({
                            'type': 'move',
                            'x': int(pt.get('x', '0')),
                            'y': int(pt.get('y', '0'))
                        })
                
                # cubicBezTo contains control and end points
                for bezier in stroke_path.findall('a:cubicBezTo', NAMESPACES):
                    pts = bezier.findall('a:pt', NAMESPACES)
                    if len(pts) >= 3:
                        # First two are control points, third is end point
                        points.append({
                            'type': 'cubic',
                            'cp1_x': int(pts[0].get('x', '0')),
                            'cp1_y': int(pts[0].get('y', '0')),
                            'cp2_x': int(pts[1].get('x', '0')),
                            'cp2_y': int(pts[1].get('y', '0')),
                            'x': int(pts[2].get('x', '0')),
                            'y': int(pts[2].get('y', '0'))
                        })
                
                if len(points) < 2:
                    continue
                
                # Get shape layout for coordinate conversion
                layout_x = shape_layout.get('x', 0)
                layout_y = shape_layout.get('y', 0)
                layout_w = shape_layout.get('width', 100)
                layout_h = shape_layout.get('height', 50)
                
                # Scale path coordinates to canvas coordinates
                def scale_x(px):
                    scaled = (px / path_w) * layout_w
                    if flipH:
                        scaled = layout_w - scaled
                    return round(layout_x + scaled, 1)
                
                def scale_y(py):
                    scaled = (py / path_h) * layout_h
                    if flipV:
                        scaled = layout_h - scaled
                    return round(layout_y + scaled, 1)
                
                # Get start and end points
                start_point = points[0]
                from_x = scale_x(start_point.get('x', 0))
                from_y = scale_y(start_point.get('y', 0))
                
                # Find the last endpoint
                end_point = None
                for p in reversed(points):
                    if p['type'] == 'cubic':
                        end_point = p
                        break
                
                if end_point is None:
                    continue
                
                to_x = scale_x(end_point.get('x', 0))
                to_y = scale_y(end_point.get('y', 0))
                
                # Calculate curve amount based on control points
                # For a typical arc, we want the midpoint's vertical offset
                # Estimate from the layout height and whether it curves up or down
                mid_y = (from_y + to_y) / 2
                
                # Check if the control points curve up or down
                # A typical arc has control points either above or below the endpoints
                curve_direction = -1  # default: curve up
                if len(points) > 1 and points[1]['type'] == 'cubic':
                    cp1_y = scale_y(points[1].get('cp1_y', 0))
                    # If control point is below the endpoints, curve is down
                    if cp1_y > mid_y:
                        curve_direction = 1
                
                # Curve amount is approximately the height of the arc
                curve_amount = layout_h * curve_direction * 0.8
                
                return {
                    'from': {'x': from_x, 'y': from_y},
                    'to': {'x': to_x, 'y': to_y},
                    'curve': round(curve_amount, 1),
                    'flip': flipH
                }
                
    except Exception as e:
        # Silently fail for shapes without arc data
        pass
    
    return None

def extract_shape_visual_data(shape, z_index, slide_xml_content=None, shape_id=None, slide_width=None):
    """Extract all visual data for a shape.
    
    Args:
        shape: The pptx shape object
        z_index: The z-index of the shape
        slide_xml_content: Optional slide XML for extracting arrow/arc data
        shape_id: Optional shape ID for XML lookups
        slide_width: Source slide width for auto-scaling coordinates
    """
    visual_data = {
        'z_index': z_index,
        'shape_type': get_shape_type_name(shape)
    }
    
    # Layout (position, size, rotation) - auto-scaled to target canvas
    layout = extract_shape_layout(shape, slide_width)
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
    font = extract_font_style(shape, slide_width)
    if font:
        visual_data['font'] = font
    
    # Connector path (for lines/arrows)
    # Check shape_type OR shape.name containing "Line" (some lines are auto_shape type)
    is_line_shape = visual_data['shape_type'] in ('connector', 'line')
    try:
        if hasattr(shape, 'name') and shape.name and 'line' in shape.name.lower():
            is_line_shape = True
    except:
        pass
    
    if is_line_shape:
        path = extract_connector_path(shape)
        if path:
            visual_data['path'] = path
        
        # Calculate actual line endpoints from layout + rotation
        # Note: layout is already scaled, so no need to pass slide_width
        if layout:
            line_endpoints = calculate_line_endpoints(layout)
            if line_endpoints:
                visual_data['line_endpoints'] = line_endpoints
    
    # Arrow head/tail ends from XML (for lines that are actually arrows)
    if slide_xml_content and shape_id:
        arrow_ends = extract_arrow_ends_from_xml(slide_xml_content, shape_id)
        if arrow_ends:
            visual_data['arrow_ends'] = arrow_ends
    
    # Arc path (for freeform arcs)
    # Note: layout is already scaled, so no need to pass slide_width
    if visual_data['shape_type'] == 'freeform' and slide_xml_content and shape_id and layout:
        arc_path = extract_arc_path_from_xml(slide_xml_content, shape_id, layout)
        if arc_path:
            visual_data['arc_path'] = arc_path
    
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

def get_group_child_ids(slide_xml_content, group_id):
    """Get all child shape IDs from a group shape.
    
    Args:
        slide_xml_content: The slide XML content
        group_id: The group shape ID
    
    Returns:
        List of child shape IDs, or None if not a group
    """
    try:
        root = ET.fromstring(slide_xml_content)
        
        # Find the group shape with matching ID
        for grpSp in root.findall('.//p:grpSp', NAMESPACES):
            cNvPr = grpSp.find('.//p:cNvPr', NAMESPACES)
            if cNvPr is not None and cNvPr.get('id') == str(group_id):
                # Found the group - get all child shape IDs
                child_ids = []
                # Get direct child shapes (sp elements)
                for sp in grpSp.findall('./p:sp', NAMESPACES):
                    child_cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
                    if child_cNvPr is not None:
                        child_ids.append(child_cNvPr.get('id'))
                # Also get child connectors (cxnSp elements)
                for cxnSp in grpSp.findall('./p:cxnSp', NAMESPACES):
                    child_cNvPr = cxnSp.find('.//p:cNvPr', NAMESPACES)
                    if child_cNvPr is not None:
                        child_ids.append(child_cNvPr.get('id'))
                return child_ids if child_ids else None
    except:
        pass
    return None

def parse_animation_sequence(slide_xml_content):
    """Parse animation sequence from slide XML and return ordered list of animation entries.
    
    Each entry contains:
    - shape_id: The shape's ID (or list of IDs if it's a group)
    - is_group: Boolean indicating if this is a group animation
    - timing: 'click' (On Click), 'with' (With Previous), or 'after' (After Previous)
    - delay: Delay in milliseconds (for 'after' timing)
    
    The animation structure in PowerPoint XML is:
    - p:timing > p:tnLst > p:par > p:cTn[nodeType="tmRoot"]
    - Inside that: p:seq > p:cTn[nodeType="mainSeq"] > p:childTnLst
    - Each click group: p:par > p:cTn[nodeType="clickPar"]
    - Inside clickPar: p:par > p:cTn[nodeType="withGroup"] or p:cTn[nodeType="afterGroup"]
    - Individual animations: p:cTn[nodeType="clickEffect"|"withEffect"|"afterEffect"]
    - Shape target: p:spTgt spid="..."
    """
    root = ET.fromstring(slide_xml_content)
    animation_entries = []
    seen_shapes = set()
    
    # Find the main sequence
    main_seq = root.find('.//p:cTn[@nodeType="mainSeq"]', NAMESPACES)
    if main_seq is None:
        # Fallback: return empty list if no animations
        return []
    
    # Get the childTnLst which contains click groups
    child_list = main_seq.find('p:childTnLst', NAMESPACES)
    if child_list is None:
        return []
    
    # Iterate through click groups (each p:par with clickPar)
    for click_group in child_list.findall('p:par', NAMESPACES):
        # Process all animations within this click group
        _process_animation_group(click_group, animation_entries, seen_shapes, 0)
    
    return animation_entries


def _process_animation_group(group_elem, entries, seen_shapes, parent_delay):
    """Recursively process animation groups to extract shape timing info.
    
    Args:
        group_elem: The p:par element to process
        entries: List to append animation entries to
        seen_shapes: Set of already-seen shape IDs (to avoid duplicates)
        parent_delay: Accumulated delay from parent afterGroup elements (in ms)
    """
    cTn = group_elem.find('p:cTn', NAMESPACES)
    if cTn is None:
        return
    
    node_type = cTn.get('nodeType', '')
    
    # Get delay from this element's stCondLst if present
    local_delay = 0
    stCondLst = cTn.find('p:stCondLst', NAMESPACES)
    if stCondLst is not None:
        cond = stCondLst.find('p:cond', NAMESPACES)
        if cond is not None:
            delay_str = cond.get('delay', '')
            if delay_str and delay_str != 'indefinite':
                try:
                    local_delay = int(delay_str)
                except ValueError:
                    pass
    
    # Determine timing type from nodeType
    timing = None
    if node_type == 'clickEffect':
        timing = 'click'
    elif node_type == 'withEffect':
        timing = 'with'
    elif node_type == 'afterEffect':
        timing = 'after'
    
    # If this is an animation effect, find the target shape
    if timing is not None:
        # Look for spTgt inside this element
        for spTgt in cTn.findall('.//p:spTgt', NAMESPACES):
            spid = spTgt.get('spid')
            if spid and spid not in seen_shapes:
                seen_shapes.add(spid)
                entry = {
                    'shape_id': spid,
                    'timing': timing
                }
                # Add delay for afterEffect (parent_delay from afterGroup + any local delay)
                total_delay = parent_delay + local_delay
                if timing == 'after' or total_delay > 0:
                    entry['delay'] = total_delay
                entries.append(entry)
                break  # One shape per animation effect
    
    # For afterGroup, accumulate the delay for child animations
    accumulated_delay = parent_delay
    if node_type == 'afterGroup':
        accumulated_delay = parent_delay + local_delay
    
    # Recursively process child elements
    child_list = cTn.find('p:childTnLst', NAMESPACES)
    if child_list is not None:
        for child_par in child_list.findall('p:par', NAMESPACES):
            _process_animation_group(child_par, entries, seen_shapes, accumulated_delay)

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
    
    # Calculate scale factor for coordinate conversion
    scale_factor = TARGET_CANVAS_WIDTH / slide_width if slide_width else 1.0
    
    presentation_data = {
        "file_path": str(file_path),
        "file_name": Path(file_path).name,
        "total_slides": len(prs.slides),
        "source_dimensions": {
            "width": slide_width,
            "height": slide_height
        },
        "target_canvas": {
            "width": TARGET_CANVAS_WIDTH,
            "height": TARGET_CANVAS_HEIGHT
        },
        "scale_factor": round(scale_factor, 3),
        "custom_shows": custom_shows,
        "slides": []
    }
    
    # Build a mapping of shape ID to pptx shape object for visual data
    with zipfile.ZipFile(file_path, 'r') as zf:
        for slide_num, slide in enumerate(prs.slides, 1):
            slide_file = f'ppt/slides/slide{slide_num}.xml'
            
            # Build shape ID to pptx shape mapping for this slide (including grouped shapes)
            pptx_shapes_by_id = {}
            for z_idx, shape, group_id in enumerate_shapes_recursive(slide.shapes):
                pptx_shapes_by_id[str(shape.shape_id)] = (shape, z_idx, group_id)
            
            try:
                slide_xml = zf.read(slide_file).decode('utf-8')
                
                # Get animation entries with timing info
                animation_entries = parse_animation_sequence(slide_xml)
                
                # Get all shapes
                shapes = parse_shapes_from_slide(slide_xml)
                
                # Build ordered animation list
                animation_sequence = []
                sequence_num = 1
                
                for anim_entry in animation_entries:
                    shape_id = anim_entry['shape_id']
                    
                    # Check if this shape_id is a group
                    child_ids = get_group_child_ids(slide_xml, shape_id)
                    
                    if child_ids:
                        # This is a group - add all child shapes with the same sequence/timing
                        for child_id in child_ids:
                            if child_id in shapes:
                                shape = shapes[child_id]
                                entry = {
                                    'sequence': sequence_num,
                                    'text': shape['text'] if shape['text'] else '',
                                    'shape_name': shape['name']
                                }
                                
                                # Add timing info (all children get same timing)
                                entry['timing'] = anim_entry['timing']
                                if 'delay' in anim_entry and anim_entry['delay'] > 0:
                                    entry['delay'] = anim_entry['delay']
                                
                                # Add visual data if available (with auto-scaling)
                                if child_id in pptx_shapes_by_id:
                                    pptx_shape, z_idx, group_id = pptx_shapes_by_id[child_id]
                                    visual = extract_shape_visual_data(pptx_shape, z_idx, slide_xml, child_id, slide_width)
                                    if visual:
                                        for key, value in visual.items():
                                            entry[key] = value
                                    # Mark that this is part of an animated group
                                    entry['group_id'] = shape_id
                                
                                # Add hyperlink info if present
                                if shape['hyperlink']:
                                    entry['hyperlink'] = shape['hyperlink']
                                    if shape['hyperlink']['type'] == 'customshow':
                                        cs_id = shape['hyperlink']['id']
                                        if cs_id in custom_shows:
                                            entry['linked_content'] = custom_shows[cs_id]
                                
                                animation_sequence.append(entry)
                        sequence_num += 1
                    elif shape_id in shapes:
                        # Regular individual shape
                        shape = shapes[shape_id]
                        # Include all animated shapes (text or not - could be rectangles, decorative shapes)
                        entry = {
                            'sequence': sequence_num,
                            'text': shape['text'] if shape['text'] else '',
                            'shape_name': shape['name']
                        }
                        
                        # Add timing info (click, with, after)
                        entry['timing'] = anim_entry['timing']
                        if 'delay' in anim_entry and anim_entry['delay'] > 0:
                            entry['delay'] = anim_entry['delay']
                        
                        # Add visual data if available (with auto-scaling)
                        if shape_id in pptx_shapes_by_id:
                            pptx_shape, z_idx, group_id = pptx_shapes_by_id[shape_id]
                            visual = extract_shape_visual_data(pptx_shape, z_idx, slide_xml, shape_id, slide_width)
                            if visual:
                                for key, value in visual.items():
                                    entry[key] = value
                            # Add group_id for debugging if shape is in a group
                            if group_id:
                                entry['group_id'] = group_id
                        
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
                # Build a set of ALL animated shape IDs (including those in animated groups)
                static_shapes = []
                animated_ids = set(e['shape_id'] for e in animation_entries)
                # Also add child IDs from animated groups
                for anim_entry in animation_entries:
                    child_ids = get_group_child_ids(slide_xml, anim_entry['shape_id'])
                    if child_ids:
                        animated_ids.update(child_ids)
                
                for shape_id, shape in shapes.items():
                    if shape_id not in animated_ids:
                        # Include all shapes - text, connectors, or decorative rectangles
                        # Skip if it's truly empty (no text, no visual importance)
                        static_entry = {
                            'text': shape['text'] if shape['text'] else '',
                            'shape_name': shape['name'],
                            'static': True
                        }
                        
                        # Add visual data if available (with auto-scaling)
                        if shape_id in pptx_shapes_by_id:
                            pptx_shape, z_idx, group_id = pptx_shapes_by_id[shape_id]
                            visual = extract_shape_visual_data(pptx_shape, z_idx, slide_xml, shape_id, slide_width)
                            if visual:
                                for key, value in visual.items():
                                    static_entry[key] = value
                            # Add group_id for debugging if shape is in a group
                            if group_id:
                                static_entry['group_id'] = group_id
                        else:
                            # Fallback: Extract visual data from XML for shapes not in pptx enumeration
                            # (e.g., shapes in groups that python-pptx doesn't expose)
                            try:
                                root = ET.fromstring(slide_xml)
                                # Find shape element by ID
                                for sp in root.findall('.//p:sp', NAMESPACES):
                                    cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
                                    if cNvPr is not None and cNvPr.get('id') == shape_id:
                                        # Extract visual data from XML
                                        visual = extract_visual_data_from_xml(sp, 0, slide_width)
                                        if visual:
                                            for key, value in visual.items():
                                                static_entry[key] = value
                                        # Check if shape is in a group by looking for parent grpSp
                                        parent = sp
                                        for _ in range(5):  # Check up to 5 levels
                                            parent = parent.find('..')
                                            if parent is None:
                                                break
                                            if parent.tag.endswith('grpSp'):
                                                grp_cNvPr = parent.find('.//p:cNvPr', NAMESPACES)
                                                if grp_cNvPr is not None:
                                                    static_entry['group_id'] = grp_cNvPr.get('id')
                                                break
                                        break
                            except:
                                pass
                        
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
    
    # Save to JSON file in extracted/ folder
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    extracted_dir = script_dir / 'extracted'
    extracted_dir.mkdir(exist_ok=True)
    
    output_filename = Path(file_path).stem + '.json'
    output_path = extracted_dir / output_filename
    
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