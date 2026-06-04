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

    left, right = st.columns([1, 1])

    with left:
        st.subheader("Canvas")
        canvas_keys = [""] + sorted(CANVAS_PRESETS)
        canvas_size = st.selectbox(
            "Paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
        )
        orientation = st.selectbox("Orientation", options=sorted(ORIENTATIONS), index=0)
        fit = st.selectbox("Image fit", options=sorted(FIT_MODES), index=0)
        background_color = st.color_picker("Background", value="#FFFFFF").lstrip("#")

    with right:
        st.subheader("Digital Layer")
        st.markdown("**Excel Mode**")
        use_custom_resolution = st.checkbox(
            "Use exact resolution",
            value=True,
            help="Set the exact number of Excel columns and rows in the generated canvas.",
        )
        if use_custom_resolution:
            width_cells = st.number_input(
                "Width cells",
                min_value=1,
                max_value=2000,
                value=240,
                step=10,
                help="Number of Excel columns across the canvas.",
            )
            height_cells = st.number_input(
                "Height cells",
                min_value=1,
                max_value=2000,
                value=170,
                step=10,
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
                help="Sets the longest canvas side in cells while preserving the selected aspect ratio.",
            )
            resolution = None

        color_count = st.slider(
            "Indexed colors",
            min_value=2,
            max_value=256,
            value=48,
            help="Maximum number of numbered colors in the Color Index.",
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
    st.subheader("Physical Layer")
    st.caption("Optional production settings for turning the Digital Layer into printable physical materials.")
    st.markdown("**Print Mode**")
    print_mode = st.checkbox(
        "Enable Print Mode",
        value=False,
        help="Adds poster splitting and optional printable color-mask sheets to the workbook.",
    )
    if print_mode:
        print_left, print_right = st.columns([1, 1])
        with print_left:
            st.markdown("**Poster splitting**")
            poster_pages_wide = st.number_input(
                "Pages wide",
                min_value=1,
                max_value=20,
                value=1,
                step=1,
                help="Number of printed pages across the poster.",
            )
            poster_pages_tall = st.number_input(
                "Pages tall",
                min_value=1,
                max_value=20,
                value=1,
                step=1,
                help="Number of printed pages from top to bottom.",
            )
        with print_right:
            st.markdown("**Color masks**")
            generate_color_masks = st.checkbox(
                "Generate color masks",
                value=False,
                help="Creates printable black-and-white mask sheets showing where each indexed color is used.",
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

    if st.button("Generate workbook", type="primary", disabled=uploaded_file is None):
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
                print_mode=print_mode,
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
    print_mode: bool,
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
            print_mode=print_mode,
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
