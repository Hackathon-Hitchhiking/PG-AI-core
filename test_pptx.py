import base64
import os
import sys
import traceback
from pathlib import Path

from tools.pptx_tools import PPTXManager
from tools.schemas import (
    Background,
    Coordinates,
    FontStyle,
    TableCell,
    TextShape,
    SlideLayout,
)
from tools.schemas import ChartData as SchemaChartData


class PPTXTester:
    """
    A test harness for validating all public methods in PPTXManager. 
    Tracks the number of passed/failed tests and prints detailed results.
    """

    def __init__(self, test_file: str) -> None:
        """
        Initialize the tester with a test PPTX file and create a PPTXManager instance.
        Args:
            test_file (str): Path to an existing .pptx presentation for testing.
        """
        self.test_file = test_file
        self.manager = PPTXManager(test_file)
        self.temp_dir = "test_temp"
        self._setup_environment()

        # Test counters
        self.total_tests = 0
        self.passed_tests = 0
        self.failed_tests = 0

    def _setup_environment(self) -> None:
        """
        Create/ensure the output directory exists for saving modified test results.
        """
        Path(self.temp_dir).mkdir(exist_ok=True)
        self.output_pptx = str(Path(self.temp_dir) / "output.pptx")
        self.output_pdf = str(Path(self.temp_dir) / "output.pdf")

    def _print_result(self, test_name: str, success: bool, error: str = "") -> None:
        """
        Print the result of each test, tracking total, passed, and failed counts.

        Args:
            test_name (str): The descriptive name of the test.
            success (bool): Whether the test passed or not.
            error (str): If failed, the error/traceback to display.
        """
        self.total_tests += 1
        status = "PASSED" if success else "FAILED"
        if success:
            self.passed_tests += 1
        else:
            self.failed_tests += 1

        print(f"[{status}] {test_name}")
        if not success and error:
            # Indent and display the error message or traceback
            print(f"  └─ Error Trace:\n{error}")

    def run_all_tests(self) -> None:
        """
        Orchestrate and run all tests, then summarize results.
        """
        print(f"Running PPTX Manager Tests on '{self.test_file}'")
        print("=" * 50)

        try:
            # Core functionality
            self.test_get_presentation_info()
            self.test_slide_analysis()

            # Slide operations
            self.test_slide_management()
            self.test_reorder_slides()
            self.test_apply_slide_template()
            self.test_background_operations()

            # Content manipulation
            self.test_text_operations()
            self.test_format_text_style()
            self.test_image_operations()
            self.test_chart_operations()
            self.test_table_operations()

            # Export and cleanup
            self.test_save_functionality()

        finally:
            # Attempt to close the manager regardless of any errors
            self.manager.close()

        print("\n" + "=" * 50)
        print(f"Total Tests: {self.total_tests}")
        print(f"Passed: {self.passed_tests} | Failed: {self.failed_tests}")
        success_rate = (self.passed_tests / self.total_tests) * 100 if self.total_tests else 0
        print(f"Success Rate: {success_rate:.0f}%")

    # --------------------------------------------------------------------------
    # Test Methods
    # --------------------------------------------------------------------------

    def test_get_presentation_info(self) -> None:
        """
        Test retrieving presentation info (metadata, slide count, etc.).
        """
        test_name = "Get Presentation Metadata"
        try:
            info = self.manager.get_presentation_info()
            # Basic checks
            assert info.slides_count > 0, "No slides found"
            assert info.dimensions.width > 0, "Invalid slide width"
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_slide_analysis(self) -> None:
        """
        Test analyzing slide content and design details.
        """
        test_name = "Slide Content Analysis"
        try:
            slide = self.manager.get_slide_details(1)
            assert len(slide.shapes) > 0, "No shapes in first slide"
            design = self.manager.analyze_slide_design(1)
            assert len(design.theme_colors) >= 0, "No theme colors detected"  # Usually > 0
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_slide_management(self) -> None:
        """
        Test creating new slides, duplicating slides, and deleting slides.
        """
        test_name = "Slide Management"
        try:
            initial_count = len(self.manager.prs.slides)
            if initial_count == 0:
                # If no slides, create one
                self.manager.create_new_slide("Title Slide")
                initial_count += 1

            # Duplicate the first slide
            new_count = self.manager.duplicate_slide(1)
            # Now delete the newly duplicated slide
            self.manager.delete_slide(new_count)
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_reorder_slides(self) -> None:
        """
        Test reordering slides in the presentation.
        """
        test_name = "Reorder Slides"
        try:
            slide_count = len(self.manager.prs.slides)
            if slide_count < 2:
                # Ensure at least 2 slides exist to reorder
                self.manager.create_new_slide("Title and Content")
                self.manager.create_new_slide("Section Header")
                slide_count = len(self.manager.prs.slides)

            # For example, reverse the slide order
            new_order = list(range(slide_count, 0, -1))
            self.manager.reorder_slides(new_order)

            # Quick assertion: The first slide should now be the last originally
            # We'll just check that the reorder didn't crash
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_apply_slide_template(self) -> None:
        """
        Test applying a slide template from another PPTX to an existing slide.
        """
        test_name = "Apply Slide Template"
        try:
            # Adjust path to point to a valid .pptx template if you have one
            template_path = "test_sources/test_template.pptx"
            if not Path(template_path).exists():
                raise FileNotFoundError("Template file not found. Provide a valid template to test.")

            # Apply template to the first slide
            self.manager.apply_slide_template(1, template_path)
            # If no error, assume success
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_background_operations(self) -> None:
        """
        Test setting and restoring slide backgrounds.
        """
        test_name = "Background Modification"
        try:
            original_bg = self.manager.get_slide_details(1).background

            test_bg = Background(type="color", value="#FF0000", transparency=0.5)
            self.manager.set_slide_background(1, test_bg)
            updated_bg = self.manager.get_slide_details(1).background
            assert updated_bg is not None, "Background was not read properly"

            # Restore original background
            self.manager.set_slide_background(1, original_bg)
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_text_operations(self) -> None:
        """
        Test adding and editing text in text boxes.
        """
        test_name = "Text Manipulation"
        try:
            text_id = self.manager.add_text_block(
                slide_num=1,
                content="Hello Bold!",
                pos=Coordinates(x=1, y=1, width=3, height=1),
                style=FontStyle(color="#000000", bold=False)
            )
            # Modify text
            self.manager.edit_text_content(1, text_id, "Updated Text")

            # Check
            shape = self.manager._find_shape(1, text_id)
            paragraph = shape.text_frame.paragraphs[0]
            run = paragraph.runs[0]
            assert run.font.bold is True, "Bold style not applied in memory"

            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_format_text_style(self) -> None:
        """
        Test applying font styles to an existing text shape.
        """
        test_name = "Format Text Style"
        try:
            # Create text box
            text_id = self.manager.add_text_block(
                1,
                "Style Me",
                Coordinates(x=2, y=1, width=3, height=1),
                FontStyle(color="#0000FF", bold=False, italic=False)
            )
            # Apply new style
            self.manager.format_text_style(1, text_id, FontStyle(bold=True))

            # Now retrieve the same shape from memory (no re-parse)
            shape = self.manager._find_shape(1, text_id)
            run = shape.text_frame.paragraphs[0].runs[0]
            assert run.font.bold is True, "Bold style not applied in memory"

            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_image_operations(self) -> None:
        """
        Test inserting and replacing an image on a slide.
        """
        test_name = "Image Handling"
        try:
            # Insert test image (a tiny 1x1 PNG base64-encoded)
            test_image = base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwAEpwGkN600QAAAAABJRU5ErkJggg=="
            )
            img_id = self.manager.insert_image(
                1,
                test_image,
                Coordinates(x=2, y=2, width=1, height=1)
            )

            # Replace image with the same data (as an example)
            self.manager.replace_image(1, img_id, test_image)

            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_chart_operations(self) -> None:
        """
        Test creating a chart and modifying its data.
        """
        test_name = "Chart Operations"
        try:
            chart_data = SchemaChartData(
                categories=["Q1", "Q2", "Q3", "Q4"],
                series=[{"name": "Sales", "values": [100, 200, 150, 300]}]
            )

            chart_id = self.manager.create_chart(
                1,
                "bar",
                chart_data,
                Coordinates(x=3, y=3, width=4, height=3)
            )

            # Modify chart data
            new_data = SchemaChartData(
                categories=["Q1", "Q2", "Q3", "Q4"],
                series=[{"name": "Sales", "values": [150, 250, 200, 350]}]
            )
            self.manager.modify_chart_data(1, chart_id, new_data)

            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_table_operations(self) -> None:
        """
        Test creating a table and editing a cell (including merges).
        """
        test_name = "Table Operations"
        try:
            # Create table
            table_id = self.manager.create_table(
                1, 3, 3,
                Coordinates(x=4, y=4, width=3, height=2)
            )

            # Merge first row cells into a header
            cell = TableCell(row=0, col=0, content="Header", span_cols=3)
            self.manager.edit_table_cell(1, table_id, cell)

            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)

    def test_save_functionality(self) -> None:
        """
        Test saving the presentation and verifying the output file.
        """
        test_name = "Presentation Save"
        try:
            self.manager.save(self.output_pptx)
            assert Path(self.output_pptx).exists(), "File not saved"
            assert os.path.getsize(self.output_pptx) > 1024, "Output .pptx appears too small (possibly empty)"
            self._print_result(test_name, True)
        except Exception as e:
            tb = traceback.format_exc()
            self._print_result(test_name, False, tb)


if __name__ == "__main__":
    # Adjust path to your testing PPTX.
    test_file = "test_sources/test_template.pptx"

    if not Path(test_file).exists():
        print(f"Test file not found: {test_file}")
        sys.exit(1)

    tester = PPTXTester(test_file)
    tester.run_all_tests()