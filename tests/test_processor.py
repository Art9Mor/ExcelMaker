from src.core.processor import process_file


def test_process_file_saves_result(sample_workbook_path):
    result = process_file(sample_workbook_path)

    assert result.success is True
    assert result.doc is not None
    assert result.output_path == sample_workbook_path
    assert result.doc.total_items == 3
    assert len(result.doc.sections) == 2
    