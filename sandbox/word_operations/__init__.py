"""Word operations module for handling various Word document operations."""

from .replace import replace_text_operation, replace_paragraph_operation
from .insert import insert_text_operation, insert_heading_operation
from .delete import delete_paragraph_operation

__all__ = [
    'read_word_content',
    'replace_text_operation',
    'replace_paragraph_operation',
    'insert_text_operation',
    'insert_heading_operation',
    'delete_paragraph_operation',
]
