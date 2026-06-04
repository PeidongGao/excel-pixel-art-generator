"""Streamlit upload/download interface for Excel pixel art generation."""

from __future__ import annotations

import tempfile
import sys
from pathlib import Path

import streamlit as st

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_pixel_art.canvas import CANVAS_PRESETS, FIT_MODES, ORIENTATIONS
from excel_pixel_art.converter import image_to_excel


def main() -> None:
    st.set_page_config(
        page_title="Excel Pixel Art Generator",
        page_icon="",
        layout="wide",
    )

    st.title("Excel Pixel Art Generator")
    st.caption("Upload an image, configure the Digital Layer, and download a paint-by-number workbook.")
    st.info(
        "**Digital Layer** controls how the image is represented as Excel cells: "
        "canvas shape, resolution, indexed colors, and cell display size."
    )

    uploaded_file = st.file_uploader(
        "Image",
        type=["png", "jpg", "jpeg", "webp", "bmp"],
        accept_multiple_files=False,
    )

    st.subheader("Canvas")
    canvas_left, canvas_right = st.columns([1, 1])
    with canvas_left:
        canvas_keys = [""] + sorted(CANVAS_PRESETS)
        canvas_size = st.selectbox(
            "Paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
        )
        orientation = st.selectbox("Orientation", options=sorted(ORIENTATIONS), index=0)
    with canvas_right:
        fit = st.selectbox("Image fit", options=sorted(FIT_MODES), index=0)
        background_color = st.color_picker("Background", value="#FFFFFF").lstrip("#")

    st.divider()
    st.subheader("Digital Layer")
    st.markdown("**Excel Mode**")
    st.caption("Defines the shared digital canvas used by both Excel and Physical outputs.")
    use_custom_resolution = st.checkbox(
        "Use exact resolution",
        value=True,
        help="Set the exact number of Excel columns and rows in the generated canvas.",
    )
    digital_left, digital_right = st.columns([1, 1])
    with digital_left:
        if use_custom_resolution:
            width_cells = st.number_input(
                "Width cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                help="Number of Excel columns across the shared canvas.",
            )
            height_cells = st.number_input(
                "Height cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                help="Number of Excel rows from top to bottom.",
            )
            resolution = (int(width_cells), int(height_cells))
            max_size = 128
        else:
            max_size = st.number_input(
                "Longest side (max cells)",
                min_value=1,
                max_value=512,
                value=128,
                step=8,
                help="Sets the longest shared canvas side while preserving the selected aspect ratio.",
            )
            resolution = None
    with digital_right:
        color_count = st.slider(
            "Indexed colors",
            min_value=2,
            max_value=256,
            value=48,
            help="Maximum number of numbered colors in the Excel Color Index.",
        )
        cell_size = st.number_input(
            "Excel cell display size",
            min_value=0.1,
            max_value=20.0,
            value=3.0,
            step=0.1,
            help="Changes how large each cell appears in Excel. It does not change resolution or color detail.",
        )

    st.divider()
    st.subheader("Outputs")
    include_excel_output = st.checkbox(
        "Excel Workbook",
        value=True,
        help="Includes the Digital Layer Reference, Template, and Color Index sheets.",
    )
    physical_output = st.checkbox(
        "Physical Output",
        value=False,
        help="Includes Print Mode sheets using the Digital Layer's indexed colors and numbers.",
    )
    if physical_output:
        st.markdown("### Physical Layer")
        st.markdown("**Print Mode**")
        st.caption(
            "Uses the same paper size, resolution, indexed colors, and color numbers as the Digital Layer."
        )
        print_left, print_right = st.columns([1, 1])
        with print_left:
            split_poster = st.checkbox(
                "Poster Split",
                value=True,
                help="Tile the shared canvas across several printed sheets when one page is too small.",
            )
            if split_poster:
                poster_pages_wide = st.number_input(
                    "Pages wide",
                    min_value=1,
                    max_value=20,
                    value=2,
                    step=1,
                    help="Number of printed pages across the poster.",
                )
                poster_pages_tall = st.number_input(
                    "Pages tall",
                    min_value=1,
                    max_value=20,
                    value=2,
                    step=1,
                    help="Number of printed pages from top to bottom.",
                )
            else:
                poster_pages_wide = 1
                poster_pages_tall = 1
        with print_right:
            generate_color_masks = st.checkbox(
                "Color Masks",
                value=False,
                help="Creates masks whose numbers match both the Color Index and Material Palette.",
            )
            st.checkbox(
                "Material Palette",
                value=True,
                disabled=True,
                help="Always included. Uses the Digital Layer's exact indexed colors and numbers.",
            )
            if generate_color_masks:
                max_color_masks = st.number_input(
                    "Masks to generate",
                    min_value=1,
                    max_value=int(color_count),
                    value=min(int(color_count), 16),
                    step=1,
                    help="Generates masks for the most-used indexed colors first.",
                )
            else:
                max_color_masks = None
    else:
        poster_pages_wide = 1
        poster_pages_tall = 1
        generate_color_masks = False
        max_color_masks = None

    if uploaded_file is not None:
        st.image(uploaded_file, caption=uploaded_file.name, use_container_width=True)

    if not include_excel_output and not physical_output:
        st.warning("Select at least one output: Excel Workbook or Physical Output.")

    if st.button(
        "Generate workbook",
        type="primary",
        disabled=uploaded_file is None or (not include_excel_output and not physical_output),
    ):
        if uploaded_file is None:
            st.warning("Upload an image first.")
            return

        with st.spinner("Converting image to Excel workbook..."):
            workbook_bytes, download_name = _build_workbook(
                uploaded_file=uploaded_file,
                max_size=int(max_size),
                cell_size=float(cell_size),
                color_count=int(color_count),
                canvas_size=canvas_size or None,
                resolution=resolution,
                orientation=orientation,
                fit=fit,
                background_color=background_color,
                include_excel_output=include_excel_output,
                physical_output=physical_output,
                poster_pages=(int(poster_pages_wide), int(poster_pages_tall)),
                generate_color_masks=generate_color_masks,
                max_color_masks=int(max_color_masks) if max_color_masks is not None else None,
            )

        st.success("Workbook ready.")
        st.download_button(
            "Download workbook",
            data=workbook_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def _build_workbook(
    uploaded_file,
    max_size: int,
    cell_size: float,
    color_count: int,
    canvas_size: str | None,
    resolution: tuple[int, int] | None,
    orientation: str,
    fit: str,
    background_color: str,
    include_excel_output: bool,
    physical_output: bool,
    poster_pages: tuple[int, int],
    generate_color_masks: bool,
    max_color_masks: int | None,
) -> tuple[bytes, str]:
    suffix = Path(uploaded_file.name).suffix or ".image"
    stem = _safe_stem(uploaded_file.name)

    with tempfile.TemporaryDirectory() as directory:
        directory_path = Path(directory)
        image_path = directory_path / f"upload{suffix}"
        output_path = directory_path / f"{stem}_pixel_art.xlsx"

        image_path.write_bytes(uploaded_file.getvalue())
        image_to_excel(
            image_path=image_path,
            output_path=output_path,
            max_size=max_size,
            cell_size=cell_size,
            color_count=color_count,
            canvas_size=canvas_size,
            resolution=resolution,
            orientation=orientation,
            fit=fit,
            background_color=background_color,
            include_excel_output=include_excel_output,
            physical_output=physical_output,
            poster_pages=poster_pages,
            generate_color_masks=generate_color_masks,
            max_color_masks=max_color_masks,
        )
        return output_path.read_bytes(), output_path.name


def _canvas_label(key: str) -> str:
    if key == "":
        return "Original image size"
    preset = CANVAS_PRESETS[key]
    return f"{preset.label} ({preset.width_mm:g} x {preset.height_mm:g} mm)"


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem.lower()
    normalized = "".join(character if character.isalnum() else "-" for character in stem)
    normalized = "-".join(part for part in normalized.split("-") if part)
    return normalized or "image"


if __name__ == "__main__":
    main()
