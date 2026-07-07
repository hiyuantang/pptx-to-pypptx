"""Internal helpers for the pptx-to-pypptx scripts.

These modules are imported by the top-level ``scripts/*.py`` commands and are
**not** meant to be run directly. Agents run the scripts; the scripts import
these shared internals (XML parsing, codegen, asset sync, slide metadata, and
.pptx utilities).
"""
