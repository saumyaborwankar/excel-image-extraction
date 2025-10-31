# Excel Image & Overlay Extraction to SVG

Complete solution for extracting images and shape overlays from Excel files and generating editable SVG outputs.

---

## 🎯 Problem Statement

You have Excel files with:
- **Images** embedded in cells
- **Shapes/drawings** overlaid on top of those images (circles, rectangles, text boxes, etc.)

**Requirements:**
1. Extract the images
2. Detect which shapes are overlaying which images
3. Create composite images that show the base image + overlays
4. **Keep the base image and overlays separate** so they can be extracted/edited later

---

## 📋 How It Works: Step-by-Step

### **Step 1: Read the Excel File**

```python
wb = load_workbook(file_path, data_only=True)
```

We use `openpyxl` to open `sample.xlsx` and read its contents.

---

### **Step 2: Extract Shape Information from Excel XML**

**Why XML?** Excel files (.xlsx) are actually ZIP archives containing XML files. Shape/drawing information is stored in `xl/drawings/drawing*.xml`.

```python
def extract_shapes_from_excel(file_path, sheet_name):
    with zipfile.ZipFile(file_path, "r") as zip_ref:
        # Find drawing XML files
        drawing_files = [f for f in zip_ref.namelist()
                        if 'xl/drawings/drawing' in f]
```

For each shape, we extract:
- **Name** (e.g., "Shape 3", "Shape 4")
- **Geometry type** (ellipse, rect, roundRect, triangle)
- **Position** (which cell it starts in + pixel offset)
- **Size** (width and height in EMUs - Excel Metric Units)
- **Colors**:
  - Fill color (RGB hex like `#FF0000`)
  - Fill transparency/alpha
  - Outline/stroke color
  - Outline width
- **Text content** (if the shape contains text)
- **Font size** (stored in hundredths of points, e.g., 1400 = 14pt)

**Example shape data:**
```python
{
    'name': 'Shape 4',
    'geometry': 'rect',
    'fill_color': None,  # No fill
    'outline_color': '#FF0000',  # Red border
    'outline_width': 4,
    'text': 'Saumya',
    'font_size_pt': 14.0,
    'from_col': 3, 'from_row': 1,  # Starts at column D, row 2
    'from_col_off': 123456,  # EMU offset from cell corner
    'from_row_off': 789012,
    'to_col': 5, 'to_row': 3,  # Ends at column F, row 4
    'width_emu': 952500,  # Width in EMUs
    'height_emu': 419100   # Height in EMUs
}
```

---

### **Step 3: Convert Excel Coordinates to Pixels**

Excel uses:
- **Cell-based positioning** (column D, row 2)
- **EMU offsets** (English Metric Units: 914,400 EMU = 1 inch)

We need **absolute pixel positions** to render overlays correctly.

```python
def emu_to_pixels(emu):
    return emu / 9525  # Conversion factor

def cell_to_pixels(ws, row, col):
    # Calculate pixel position by summing column widths and row heights
    x = sum(column_widths) * 7  # Excel unit conversion
    y = sum(row_heights) * 1.33
    return x, y
```

**For each shape:**
```python
# Get cell position
cell_x, cell_y = cell_to_pixels(ws, from_row + 1, from_col + 1)

# Add offset from cell corner
abs_x = cell_x + emu_to_pixels(from_col_off)
abs_y = cell_y + emu_to_pixels(from_row_off)

# Calculate size
width_px = emu_to_pixels(width_emu)
height_px = emu_to_pixels(height_emu)
```

Now we know: **"Shape 4 is at pixel (375, 140) with size 100x44"**

---

### **Step 4: Extract Images from Excel**

```python
for image in ws._images:
    # Get image position (same cell-based + EMU system)
    # Get image bytes
    img_bytes = image._data()

    # Save as PNG
    with open(img_path, 'wb') as f:
        f.write(img_bytes)

    # Load for processing
    pil_image = Image.open(io.BytesIO(img_bytes))
```

We extract:
- Image position (absolute x, y in pixels)
- Image size (calculated width/height from Excel cells)
- **Actual image dimensions** (1024x1024 pixels for example)
- Image data as PIL Image object

**Key insight:** Excel's calculated size (191x191px) ≠ actual image size (1024x1024px)

This means we need a **scaling factor**: `scale_x = 1024/191 = 5.361`

---

### **Step 5: Detect Overlays**

For each image, check if shapes overlay it:

```python
def rectangles_overlap(rect1, rect2):
    # Check if two rectangles intersect
    x1, y1, w1, h1 = rect1
    x2, y2, w2, h2 = rect2

    return not (
        x1 + w1 <= x2  # rect1 is left of rect2
        or x2 + w2 <= x1  # rect2 is left of rect1
        or y1 + h1 <= y2  # rect1 is above rect2
        or y2 + h2 <= y1  # rect2 is above rect1
    )
```

**Example:**
```
Image at (329, 7) size 191x191
Shape at (375, 140) size 100x44

Do they overlap? YES
375 is between 329 and 520 (329+191) ✓
140 is between 7 and 198 (7+191) ✓
```

---

### **Step 6: Generate SVG Output (THE KEY PART)**

This is where the magic happens! Instead of just creating a flat image, we create an **SVG file** that keeps everything separate.

#### 6a. **Embed Base Image as Base64**

```python
def image_to_base64(pil_image):
    buffered = io.BytesIO()
    pil_image.save(buffered, format="PNG")
    img_bytes = buffered.getvalue()
    return base64.b64encode(img_bytes).decode('utf-8')
```

This converts the PNG to a text string like:
```
iVBORw0KGgoAAAANSUhEUgAABAAAAAQACAMAAABIw9ux...
```

#### 6b. **Create SVG Structure**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg"
     width="1024" height="1024">

  <!-- BASE IMAGE (embedded) -->
  <image id="base-image"
         xlink:href="data:image/png;base64,iVBORw0KG..."
         width="1024" height="1024"
         x="0" y="0"/>

  <!-- OVERLAYS (editable shapes) -->
  <g id="overlays">

    <!-- Shape 1: Circle with red border -->
    <g id="Shape 3" class="shape-overlay" data-geometry="ellipse">
      <ellipse cx="503.5" cy="478.5" rx="278.5" ry="313.5"
               fill="none"
               stroke="#FF0000" stroke-width="4"/>
    </g>

    <!-- Shape 2: Text box -->
    <g id="Shape 4" class="shape-overlay" data-geometry="rect">
      <rect x="246" y="716" width="536" height="235"
            fill="none" stroke="none"/>
      <text x="514.0" y="833.5"
            font-size="75px"
            font-family="Arial, sans-serif"
            text-anchor="middle"
            fill="#000000">Saumya</text>
    </g>

  </g>
</svg>
```

#### 6c. **Apply Scaling to Shapes**

Remember: Excel calculated 191x191, actual image is 1024x1024

```python
scale_x = 1024 / 191 = 5.361
scale_y = 1024 / 191 = 5.361

# Scale shape position
rel_x = (shape_x - image_x) * scale_x
rel_y = (shape_y - image_y) * scale_y

# Scale shape size
width_scaled = shape_width * scale_x
height_scaled = shape_height * scale_y

# Scale font size
scaled_font_size = 14pt * 5.361 = 75px
```

**Result:** Shapes are positioned correctly on the high-resolution image!

---

### **Step 7: Also Create JPG Preview**

For quick viewing, we also rasterize to JPG:

```python
# Create composite using PIL
composite = base_img.convert("RGBA")
draw = ImageDraw.Draw(composite)

# Draw each shape
for shape in shape_overlays:
    if shape['geometry'] == 'ellipse':
        draw.ellipse([x, y, x2, y2], fill=fill, outline=stroke)
    elif shape['geometry'] == 'rect':
        draw.rectangle([x, y, x2, y2], fill=fill, outline=stroke)
    # ... etc

# Save as JPG
composite_rgb = composite.convert("RGB")
composite_rgb.save(path, "JPEG", quality=95)
```

---

## 📁 Output Files Created

### Main Outputs

```
images_with_overlays/
├── D1-F10_with_overlays.svg  (912 KB)
│   ├── Base image (embedded as base64)
│   └── Overlays (editable SVG shapes)
│
└── D1-F10_with_overlays.jpg  (387 KB)
    └── Flattened preview image
```

### Separated Layers (via utility scripts)

```
separated_layers/
├── D1-F10_with_overlays_base_image.png  (683 KB)
│   └── Extracted PNG (decoded from base64)
│
├── D1-F10_with_overlays_overlays_only.svg  (781 B)
│   └── Just the shapes, no base image
│
├── D1-F10_with_overlays_shape_Shape 3.svg  (438 B)
│   └── Individual circle shape
│
└── D1-F10_with_overlays_shape_Shape 4.svg  (585 B)
    └── Individual text box shape
```

---

## 🔄 How Components Stay Separate

### In the SVG:

1. **Base Image Element:**
   ```xml
   <image id="base-image" xlink:href="data:image/png;base64,..."/>
   ```
   - Self-contained (base64 encoded)
   - Can be extracted by decoding base64

2. **Shape Elements:**
   ```xml
   <ellipse cx="503.5" cy="478.5" rx="278.5" ry="313.5"
            fill="none" stroke="#FF0000"/>
   ```
   - Native SVG geometry
   - Fully editable (change colors, positions, sizes)
   - Can be removed/added

3. **Grouped Separately:**
   ```xml
   <g id="overlays">
     <!-- All shapes here -->
   </g>
   ```

---

## 💡 Why This Approach Works

### ✅ **Separation**
- Base image: `<image>` element with base64 data
- Overlays: `<ellipse>`, `<rect>`, `<text>` elements
- Can extract/edit independently

### ✅ **Editability**
- Open SVG in Inkscape/Illustrator
- Change shape colors: edit `fill="#FF0000"`
- Move shapes: edit `x`, `y`, `cx`, `cy`
- Edit text: change text content
- Add/remove shapes: add/remove `<g>` elements

### ✅ **Extractability**

**Get base image:**
```python
# Parse SVG
tree = ET.parse('file.svg')
image = tree.find('.//image[@id="base-image"]')

# Extract base64
href = image.get('{http://www.w3.org/1999/xlink}href')
base64_data = href.split(',', 1)[1]

# Decode to PNG
png_bytes = base64.b64decode(base64_data)
```

**Get shapes:**
```python
overlays = tree.find('.//g[@id="overlays"]')
shapes = overlays.findall('.//g[@class="shape-overlay"]')

for shape in shapes:
    shape_id = shape.get('id')  # "Shape 3"
    geometry = shape.get('data-geometry')  # "ellipse"
    # ... extract properties
```

### ✅ **Standards-Based**
- Uses standard SVG 1.1 spec
- Works with all SVG tools
- Can convert to other formats (PDF, PNG, etc.) while preserving structure

---

## 🎬 Complete Flow Diagram

```
Excel File (sample.xlsx)
    │
    ├─── Images (embedded PNGs)
    │     └─── Extract to memory
    │
    └─── Drawings XML (xl/drawings/drawing*.xml)
          └─── Parse to get shape data

                    ↓

Calculate Positions & Detect Overlays
    │
    ├─── Convert Excel coordinates → Pixels
    ├─── Check which shapes overlay which images
    └─── Calculate scaling factors

                    ↓

Generate SVG
    │
    ├─── Convert base image → Base64
    ├─── Create <image> element with base64
    ├─── Create <ellipse>, <rect>, <text> for each shape
    └─── Apply scaling to positions/sizes

                    ↓

Save Outputs
    │
    ├─── SVG (editable, separate components)
    └─── JPG (flattened preview)

                    ↓

Utilities (optional)
    │
    ├─── Extract base image (decode base64 → PNG)
    ├─── Extract shapes (create individual SVGs)
    └─── Create overlays-only SVG (remove base image)
```

---

## 🧪 Example: What Happens to "Shape 4"

1. **Excel:** Shape at cell D2, offset 123456 EMU, size 952500x419100 EMU
2. **Convert:** Position (375, 140) px, size (100, 44) px
3. **Detect:** Overlays image at (329, 7) size 191x191
4. **Scale:**
   - Relative position: (375-329, 140-7) = (46, 133)
   - Scaled: (46×5.361, 133×5.361) = (246, 716)
   - Scaled size: (100×5.361, 44×5.361) = (536, 235)
   - Font: 14pt × 5.361 = 75px
5. **SVG:**
   ```xml
   <rect x="246" y="716" width="536" height="235"/>
   <text x="514" y="833.5" font-size="75px">Saumya</text>
   ```
6. **Result:** Perfectly positioned text box on 1024x1024 image

---

## 🚀 Usage

### Installation

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate  # On macOS/Linux
# or
venv\Scripts\activate  # On Windows

# Install dependencies
pip install openpyxl Pillow
```

### Run Main Script

```bash
# Generate SVG and JPG files
python3 extract.py
```

**Output:**
- `images_with_positions/*.png` - Individual extracted images
- `images_with_overlays/*.svg` - SVG files with embedded images + overlays
- `images_with_overlays/*.jpg` - Rasterized preview files

### Extract Components (Optional)

```bash
# View SVG component details
python3 extract_svg_components.py

# Separate into individual layers
python3 separate_svg_layers.py
```

**Output:**
- `separated_layers/*_base_image.png` - Extracted base images
- `separated_layers/*_overlays_only.svg` - Overlays without base image
- `separated_layers/*_shape_*.svg` - Individual shape SVGs

---

## 📂 Project Structure

```
attemptv2/
├── extract.py                          # Main script
├── extract_svg_components.py       # Utility: inspect SVG structure
├── separate_svg_layers.py          # Utility: separate layers
├── sample.xlsx                     # Input Excel file
├── README.md                       # This file
├── SVG_SOLUTION.md                 # Technical documentation
│
├── images_with_positions/          # Extracted individual images
├── images_with_overlays/           # SVG + JPG composites
└── separated_layers/               # Separated components
```

---

## 🔑 Key Features

- ✅ **Separation** - Base image and overlays remain completely independent
- ✅ **Editability** - All overlays can be edited in vector editors
- ✅ **Scalability** - SVG is resolution-independent
- ✅ **Extractability** - Components can be extracted programmatically
- ✅ **Standards-based** - Uses standard SVG format, compatible with all tools
- ✅ **Self-contained** - No external dependencies (base64 embedding)
- ✅ **Correct scaling** - Shapes positioned accurately on high-res images
- ✅ **All properties preserved** - Colors, fonts, transparency, geometry

---

## 📖 Documentation

- **README.md** (this file) - Complete walkthrough and usage guide
- **SVG_SOLUTION.md** - Technical documentation and API reference

---

## 🎯 Result

You now have:
- ✅ SVG files with **embedded base images** (base64)
- ✅ **Editable shape overlays** as native SVG elements
- ✅ **Complete separation** - can extract/edit components independently
- ✅ **Correct scaling** - shapes positioned accurately on high-res images
- ✅ **All properties preserved** - colors, fonts, transparency, geometry
- ✅ **Standard format** - works with all SVG tools
- ✅ **Self-contained** - no external dependencies

The goal of **keeping base images and overlays separate even after export** is fully achieved! 🎉

---

## 📝 License

This project is provided as-is for educational and commercial use.

---

## 🤝 Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.
