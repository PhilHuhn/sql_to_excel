import os
import pytest
from file_creator import ExcelGenerator


def test_formatter():
    # Arrange
    json_file = 'order_leg.json'
    output_file = 'test_output.xlsx'

    # Act
    formatter = ExcelGenerator(json_file, test_mode=True)
    formatter.generate_excel(output_file)

    # Assert
    assert os.path.exists(output_file)

    # Clean up
    #os.remove(output_file)
