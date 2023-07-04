import requests
from tqdm import tqdm
import os
from PIL import Image
import numpy as np
from docx import Document
from docx.shared import Pt, Inches

def get_center(box):
    """
    Calculate the center point of a polygon.

    Parameters:
    - box: list representing the polygon coordinates.

    Returns:
    - Tuple (center_x, center_y) representing the center point of the polygon.
    """
    _x_list = [coord [0] for coord in box]
    _y_list = [coord [1] for coord in box]
    _len = len(box)
    _x = sum(_x_list) / _len
    _y = sum(_y_list) / _len
    return(_x, _y)
def get_size(box):
    """
    Calculate the size (width and height) of a polygon.

    Parameters:
    - box: list representing the polygon coordinates.

    Returns:
    - Tuple (width, height) representing the size of the polygon.
    """
    _x_list = [coord [0] for coord in box]
    _y_list = [coord [1] for coord in box]
    return(max(_x_list)-min(_x_list),max(_y_list)-min(_y_list))

def arrange_poly_bbox(bboxes):
    """
    Arrange the given list of bounding boxes (polygons) based on their positions.

    Parameters:
    - bboxes: List of polygonn

    Returns:
    - Numpy array representing the arrangement of polygons as a directed graph.
    """
    
    n = len(bboxes)
    
    centers=[get_center(box) for box in bboxes]
    xcentres = [center[0] for center in centers]
    ycentres = [center[1] for center in centers]

    box_sizes=[get_size(box) for box in bboxes]
    heights = [box_size[1] for box_size in box_sizes]
    width = [box_size[0] for box_size in box_sizes]

    def is_top_to(i, j):
        result = (ycentres[j] - ycentres[i]) > ((heights[i] + heights[j]) / 3)
        return result

    def is_left_to(i, j):
        return (xcentres[i] - xcentres[j]) > ((width[i] + width[j]) / 3)

    # <L-R><T-B>
    # +1: Left/Top
    # -1: Right/Bottom
    g = np.zeros((n, n), dtype='int')
    for i in range(n):
        for j in range(n):
            if is_left_to(i, j):
                g[i, j] += 10
            if is_left_to(j, i):
                g[i, j] -= 10
            if is_top_to(i, j):
                g[i, j] += 1
            if is_top_to(j, i):
                g[i, j] -= 1
    return g

def arrange_row(bboxes=None, g=None, i=None, visited=None):
    """
    Recursively arrange the rows of polygons based on the given directed graph.

    Parameters:
    - bboxes: List of numpy arrays, each representing the coordinates of a polygon.
    - g: Numpy array representing the directed graph of polygon arrangement.
    - i: Current index of the row being arranged.
    - visited: List of indices that have been visited.

    Returns:
    - List of rows, where each row is a list of polygon indices.
    """

    if visited is not None and i in visited:
        return []
    if g is None:
        g = arrange_poly_bbox(bboxes)
    if i is None:
        visited = []
        rows = []
        for i in range(g.shape[0]):
            if i not in visited:
                indices = arrange_row(g=g, i=i, visited=visited)
                visited.extend(indices)
                rows.append(indices)
        return rows
    else:
        indices = [j for j in range(g.shape[0]) if j not in visited]
        indices = [j for j in indices if abs(g[i, j]) == 10 or i == j]
        indices = np.array(indices)
        g_ = g[np.ix_(indices, indices)]
        order = np.argsort(np.sum(g_, axis=1))
        indices = indices[order].tolist()
        indices = [int(i) for i in indices]
        return indices


def row_to_textline(row, texts, polygons_A4, side_width=87, tab_width=42):
    """
    Generate a text line from a row of polygon indices, corresponding texts, and polygon coordinates.

    Parameters:
    - row: List of polygon indices representing the row.
    - texts: List of texts corresponding to each polygon.
    - polygons_A4: List of polygons' coordinates mapped to A4 size.
    - side_width: Width of the sides of the document.
    - tab_width: Width of a tab space.

    Returns:
    - Generated text line.
    """
    
    textline = ""
    first_id = row[0]
    first_space = max(min([x[0] for x in polygons_A4[first_id]]) - side_width, 0)
    textline += "\t"*int(first_space/tab_width)
    textline += texts[first_id]

    for i, id in enumerate(row):
        if i == 0:
            continue
        text_distance = min([x[0] for x in polygons_A4[id]]) - max([x[0] for x in polygons_A4[row[i-1]]])
        num_tab = int(text_distance/tab_width)
        if num_tab == 0:
            num_space = max(int(text_distance/tab_width*12), 1)
            textline += " "*num_space
        else:
            textline += "\t"*num_tab
        textline += texts[id]
    return textline

def doc_gen(texts, polygons, w, h):
    """
    Generate a document based on the given texts and polygons.

    Parameters:
    - texts: List of texts corresponding to each polygon.
    - polygons: List of polygons.
    - w: Width of the original image.
    - h: Height of the original image.

    Returns:
    - Document object (docx.Document) containing the generated document.
    """

    text_width = 7
    space_width = text_width/2
    tab_width = space_width*12
    line_width = 78*text_width
    side_width = int(line_width/6.27)
    w_A4 = line_width + 2*side_width
    h_A4 = int(w_A4/8.27*11.69)

    rows = arrange_row(bboxes= polygons)
    
    polygons_A4 = [[[int(point[0]/w*w_A4), int(point[1]/h*h_A4)] 
                for point in polygon]
                for polygon in polygons]
    
    doc = Document()

    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(12)

    section = doc.sections[0]
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    paragraph_format = doc.styles['Normal'].paragraph_format
    paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
    
    for row in rows:
        linetext = row_to_textline(row, texts, polygons_A4, side_width=side_width, tab_width=tab_width)  
        doc.add_paragraph(linetext)
    
    # doc.save(output_path)
    return doc


def docx_file_gen(image_path, output_path): 
    api_url="http://10.10.1.37:10000/ocr/"
    post_param = dict(output_features='fulltext', refine_link=True) #,link_threshold=0.7,

    filename=os.path.basename(image_path)
    img_name, ext = os.path.splitext(filename)

    image=Image.open(image_path)
    w,h=image.size
    res = requests.post(api_url, files=dict(image=open(image_path, "rb")), data=post_param)
    texts,polygons=res.json()['results'][0]['texts'],res.json()['results'][0]['boxes']
    # texts
    doc = doc_gen(texts, polygons, w, h)

    doc.save(output_path)

# image_path = "/home/phung/AnhHung/spade-rewrite-v2/temp/table_image/image/test1.png"
image_path = "./input/test3.jpg"
output_path = "./output/test3.docx"
docx_file_gen(image_path, output_path)
print("Done!")