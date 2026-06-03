# Excel Pixel Art Generator

Convert an image into an Excel workbook where each cell is colored like a pixel.

Each generated workbook includes:

- `Reference` - the finished colored pixel-art image.
- `Template` - a paper-style paint-by-number canvas with light gray color index numbers.
- `Color Index` - one row per indexed color, with number, hex color code, swatch, and cell count.

The template uses Excel screen gridlines for editing, not real cell borders, so filled artwork can print without black border lines.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e .
```

## Usage

```bash
excel-pixel-art path/to/image.png
excel-pixel-art path/to/image.png -o output.xlsx --max-size 96
```

With the included sample image:

```bash
excel-pixel-art image/UnderTheWaveOffKanagawa.jpg -o wave.xlsx --max-size 128
excel-pixel-art image/IMG_0148.JPG -o landscape.xlsx --max-size 128
```

Use a paper canvas to auto-fit the image into a fixed Excel drawing area:

```bash
excel-pixel-art image/IMG_0148.JPG -o landscape_a4.xlsx --canvas-size a4 --orientation landscape --max-size 160
excel-pixel-art image/UnderTheWaveOffKanagawa.jpg -o wave_letter.xlsx --canvas-size letter --fit contain --max-size 160
```

Set an exact Excel cell resolution when you want full manual control:

```bash
excel-pixel-art image/IMG_0148.JPG -o custom_160x100.xlsx --resolution 160x100 --fit contain
excel-pixel-art image/UnderTheWaveOffKanagawa.jpg -o poster_240x160.xlsx --canvas-size a3 --resolution 240x160 --orientation landscape
```

`--resolution WIDTHxHEIGHT` means Excel columns x rows. When it is provided, it overrides `--max-size` and uses the exact grid size you request. You can still combine it with `--canvas-size` to set Excel print paper setup.

Control how many unique indexed colors are used:

```bash
excel-pixel-art image/IMG_0148.JPG -o photo_32_colors.xlsx --resolution 160x100 --colors 32
```

`--colors` keeps the color index practical for photos by merging near-duplicate colors into a smaller palette. Default: `48`.

High-detail A4 example:

```bash
excel-pixel-art image/IMG_0148.JPG -o output/landscape_a4_high_detail_240x170_256colors.xlsx --canvas-size a4 --orientation landscape --resolution 240x170 --colors 256
```

This keeps the workbook print setup as A4 landscape while using a `240 x 170` Excel-cell template and a 256-color index.

Common canvas presets:

- `a0` - 841 x 1189 mm
- `a1` - 594 x 841 mm
- `a2` - 420 x 594 mm
- `a3` - 297 x 420 mm
- `a4` - 210 x 297 mm
- `a5` - 148 x 210 mm
- `a6` - 105 x 148 mm
- `b0` - 1000 x 1414 mm
- `b1` - 707 x 1000 mm
- `b2` - 500 x 707 mm
- `b3` - 353 x 500 mm
- `b4` - 250 x 353 mm
- `b5` - 176 x 250 mm
- `b6` - 125 x 176 mm
- `letter` - 215.9 x 279.4 mm
- `legal` - 215.9 x 355.6 mm
- `executive` - 190.5 x 254.0 mm
- `folio` - 210.0 x 330.0 mm
- `ledger` - 432.0 x 279.0 mm
- `tabloid` - 279.0 x 432.0 mm
- `square` - 210 x 210 mm

ISO and common format dimensions are based on Engineering ToolBox's paper drafting size tables.

Canvas options:

- `--resolution WIDTHxHEIGHT`
- `--colors 48`
- `--orientation auto|portrait|landscape`
- `--fit contain|cover`
- `--background-color FFFFFF`

You can also run the compatibility wrapper:

```bash
python img_to_excel.py path/to/image.png
```

## Web Upload

Start the local upload app:

```bash
excel-pixel-art-web
```

Then open http://127.0.0.1:8000, upload a painting or photo, and download the generated workbook.

## Development

Run the tests with:

```bash
python -m unittest
```
