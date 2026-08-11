"""Headless slide-rendering helpers shared by public scripts."""

import shutil
import subprocess
import tempfile
from pathlib import Path


class RenderError(RuntimeError):
    """Raised when the local rendering toolchain cannot produce a slide image."""


def render_slides(pptx_path: Path, slides: list[int], output_dir: Path, dpi: int = 150) -> dict[int, Path]:
    """Render selected 1-based slide numbers to PNG via LibreOffice and pdftoppm."""
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    if not pptx_path.exists():
        raise FileNotFoundError(f"Target PPTX not found: {pptx_path}")
    if not shutil.which("soffice"):
        raise RenderError("LibreOffice (soffice) not found")
    if not shutil.which("pdftoppm"):
        raise RenderError("pdftoppm not found")
    if dpi <= 0:
        raise ValueError("DPI must be positive")

    output_dir.mkdir(parents=True, exist_ok=True)
    rendered_paths: dict[int, Path] = {}
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        profile_dir = tmp / "libreoffice-profile"
        profile_dir.mkdir()
        result = subprocess.run(
            [
                "soffice",
                f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                "--headless",
                "--convert-to",
                'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}',
                "--outdir",
                str(tmp),
                str(pptx_path),
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).strip()
            raise RenderError(f"LibreOffice conversion failed: {detail}")

        pdf_files = list(tmp.glob("*.pdf"))
        if not pdf_files:
            raise RenderError("LibreOffice conversion produced no PDF")
        pdf_path = pdf_files[0]

        for slide_num in slides:
            prefix = tmp / f"slide_{slide_num}"
            result = subprocess.run(
                [
                    "pdftoppm",
                    "-png",
                    "-singlefile",
                    "-f",
                    str(slide_num),
                    "-l",
                    str(slide_num),
                    "-r",
                    str(dpi),
                    str(pdf_path),
                    str(prefix),
                ],
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                detail = (result.stderr or result.stdout).strip()
                raise RenderError(f"Slide {slide_num} rendering failed: {detail}")
            rendered = prefix.with_suffix(".png")
            if not rendered.exists():
                raise RenderError(f"Slide {slide_num} rendering produced no PNG")
            dest = output_dir / f"slide_{slide_num}.png"
            shutil.move(str(rendered), str(dest))
            rendered_paths[slide_num] = dest

    return rendered_paths
