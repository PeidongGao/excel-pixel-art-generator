"""Streamlit upload/download interface for Excel pixel art generation."""

from __future__ import annotations

import sys
import tempfile
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
    st.caption("Upload one image, then generate completely separate Digital and Physical outputs.")

    uploaded_file = st.file_uploader(
        "Image",
        type=["png", "jpg", "jpeg", "webp", "bmp"],
        accept_multiple_files=False,
    )
    if uploaded_file is not None:
        st.image(uploaded_file, caption=uploaded_file.name, use_container_width=True)

    st.divider()
    _digital_layer(uploaded_file)
    st.divider()
    _physical_layer(uploaded_file)


def _digital_layer(uploaded_file) -> None:
    st.header("Digital Layer")
    st.markdown("**Excel Mode**")
    st.caption("V1 Excel pixel-art workflow and workbook output.")

    left, right = st.columns([1, 1])
    with left:
        st.subheader("Canvas")
        canvas_keys = [""] + sorted(CANVAS_PRESETS)
        canvas_size = st.selectbox(
            "Paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
            key="digital_canvas_size",
        )
        orientation = st.selectbox(
            "Orientation",
            options=sorted(ORIENTATIONS),
            index=0,
            key="digital_orientation",
        )
        fit = st.selectbox(
            "Image fit",
            options=sorted(FIT_MODES),
            index=0,
            key="digital_fit",
        )
        background_color = st.color_picker(
            "Background",
            value="#FFFFFF",
            key="digital_background",
        ).lstrip("#")

    with right:
        st.subheader("Detail")
        use_custom_resolution = st.checkbox(
            "Use exact resolution",
            value=True,
            key="digital_exact_resolution",
        )
        if use_custom_resolution:
            width_cells = st.number_input(
                "Width cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                key="digital_width",
            )
            height_cells = st.number_input(
                "Height cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                key="digital_height",
            )
            resolution = (int(width_cells), int(height_cells))
            max_size = 128
        else:
            max_size = st.number_input(
                "Max cells",
                min_value=1,
                max_value=512,
                value=128,
                step=8,
                key="digital_max_size",
            )
            resolution = None

        color_count = st.slider(
            "Indexed colors",
            min_value=2,
            max_value=256,
            value=48,
            key="digital_colors",
        )
        cell_size = st.number_input(
            "Excel cell size",
            min_value=0.1,
            max_value=20.0,
            value=3.0,
            step=0.1,
            key="digital_cell_size",
        )

    if st.button(
        "Generate Digital workbook",
        type="primary",
        disabled=uploaded_file is None,
        key="generate_digital",
    ):
        with st.spinner("Generating Digital workbook..."):
            workbook_bytes, download_name = _build_digital_workbook(
                uploaded_file=uploaded_file,
                max_size=int(max_size),
                cell_size=float(cell_size),
                color_count=int(color_count),
                canvas_size=canvas_size or None,
                resolution=resolution,
                orientation=orientation,
                fit=fit,
                background_color=background_color,
            )
        st.download_button(
            "Download Digital workbook",
            data=workbook_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_digital",
        )


def _physical_layer(uploaded_file) -> None:
    st.header("Physical Layer")
    st.markdown("**Print Mode**")
    st.caption("Independent print canvas, material palette, poster split, masks, and workbook output.")

    canvas_column, detail_column, production_column = st.columns([1, 1, 1])
    canvas_keys = [""] + sorted(CANVAS_PRESETS)

    with canvas_column:
        st.subheader("Print Canvas")
        canvas_size = st.selectbox(
            "Physical paper size",
            options=canvas_keys,
            index=canvas_keys.index("a4"),
            format_func=_canvas_label,
            key="physical_canvas_size",
        )
        orientation = st.selectbox(
            "Physical orientation",
            options=sorted(ORIENTATIONS),
            index=0,
            key="physical_orientation",
        )
        fit = st.selectbox(
            "Physical image fit",
            options=sorted(FIT_MODES),
            index=0,
            key="physical_fit",
        )
        background_color = st.color_picker(
            "Physical background",
            value="#FFFFFF",
            key="physical_background",
        ).lstrip("#")

    with detail_column:
        st.subheader("Material Detail")
        use_custom_resolution = st.checkbox(
            "Use exact physical resolution",
            value=True,
            key="physical_exact_resolution",
        )
        if use_custom_resolution:
            width_cells = st.number_input(
                "Physical width cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                key="physical_width",
            )
            height_cells = st.number_input(
                "Physical height cells",
                min_value=1,
                max_value=2000,
                value=128,
                step=8,
                key="physical_height",
            )
            resolution = (int(width_cells), int(height_cells))
            max_size = 128
        else:
            max_size = st.number_input(
                "Physical max cells",
                min_value=1,
                max_value=512,
                value=128,
                step=8,
                key="physical_max_size",
            )
            resolution = None

        material_color_count = st.slider(
            "Material Palette colors",
            min_value=2,
            max_value=256,
            value=48,
            key="physical_colors",
        )
        cell_size = st.number_input(
            "Print cell size",
            min_value=0.1,
            max_value=20.0,
            value=3.0,
            step=0.1,
            key="physical_cell_size",
        )

    with production_column:
        st.subheader("Production")
        split_poster = st.checkbox("Poster Split", value=True, key="physical_poster_split")
        if split_poster:
            poster_pages_wide = st.number_input(
                "Pages wide",
                min_value=1,
                max_value=20,
                value=2,
                step=1,
                key="physical_pages_wide",
            )
            poster_pages_tall = st.number_input(
                "Pages tall",
                min_value=1,
                max_value=20,
                value=2,
                step=1,
                key="physical_pages_tall",
            )
        else:
            poster_pages_wide = 1
            poster_pages_tall = 1

        generate_color_masks = st.checkbox("Color Masks", value=False, key="physical_masks")
        if generate_color_masks:
            max_color_masks = st.number_input(
                "Masks to generate",
                min_value=1,
                max_value=int(material_color_count),
                value=min(int(material_color_count), 16),
                step=1,
                key="physical_mask_count",
            )
        else:
            max_color_masks = None
        st.checkbox("Material Palette", value=True, disabled=True, key="physical_palette")

    if st.button(
        "Generate Physical workbook",
        type="primary",
        disabled=uploaded_file is None,
        key="generate_physical",
    ):
        with st.spinner("Generating Physical workbook..."):
            workbook_bytes, download_name = _build_physical_workbook(
                uploaded_file=uploaded_file,
                max_size=int(max_size),
                cell_size=float(cell_size),
                material_color_count=int(material_color_count),
                canvas_size=canvas_size or None,
                resolution=resolution,
                orientation=orientation,
                fit=fit,
                background_color=background_color,
                poster_pages=(int(poster_pages_wide), int(poster_pages_tall)),
                generate_color_masks=generate_color_masks,
                max_color_masks=int(max_color_masks) if max_color_masks is not None else None,
            )
        st.download_button(
            "Download Physical workbook",
            data=workbook_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_physical",
        )


def _build_digital_workbook(
    uploaded_file,
    max_size: int,
    cell_size: float,
    color_count: int,
    canvas_size: str | None,
    resolution: tuple[int, int] | None,
    orientation: str,
    fit: str,
    background_color: str,
) -> tuple[bytes, str]:
    return _build_workbook(
        uploaded_file=uploaded_file,
        output_suffix="digital",
        max_size=max_size,
        cell_size=cell_size,
        color_count=color_count,
        canvas_size=canvas_size,
        resolution=resolution,
        orientation=orientation,
        fit=fit,
        background_color=background_color,
        include_excel_output=True,
        physical_output=False,
    )


def _build_physical_workbook(
    uploaded_file,
    max_size: int,
    cell_size: float,
    material_color_count: int,
    canvas_size: str | None,
    resolution: tuple[int, int] | None,
    orientation: str,
    fit: str,
    background_color: str,
    poster_pages: tuple[int, int],
    generate_color_masks: bool,
    max_color_masks: int | None,
) -> tuple[bytes, str]:
    return _build_workbook(
        uploaded_file=uploaded_file,
        output_suffix="physical",
        max_size=96,
        cell_size=3.0,
        color_count=48,
        canvas_size=None,
        resolution=None,
        orientation="auto",
        fit="contain",
        background_color="FFFFFF",
        include_excel_output=False,
        physical_output=True,
        physical_max_size=max_size,
        physical_cell_size=cell_size,
        material_color_count=material_color_count,
        physical_canvas_size=canvas_size,
        physical_resolution=resolution,
        physical_orientation=orientation,
        physical_fit=fit,
        physical_background_color=background_color,
        poster_pages=poster_pages,
        generate_color_masks=generate_color_masks,
        max_color_masks=max_color_masks,
    )


def _build_workbook(uploaded_file, output_suffix: str, **options) -> tuple[bytes, str]:
    suffix = Path(uploaded_file.name).suffix or ".image"
    stem = _safe_stem(uploaded_file.name)

    with tempfile.TemporaryDirectory() as directory:
        directory_path = Path(directory)
        image_path = directory_path / f"upload{suffix}"
        output_path = directory_path / f"{stem}_{output_suffix}.xlsx"

        image_path.write_bytes(uploaded_file.getvalue())
        image_to_excel(image_path=image_path, output_path=output_path, **options)
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
