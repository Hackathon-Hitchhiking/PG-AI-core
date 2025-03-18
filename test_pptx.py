import base64
import os
import sys

from pathlib import Path

from tools.pptx_tools import PPTXManager
from tools.schemas import Background, Coordinates, FontStyle, TableCell, TextShape
from tools.schemas import ChartData as SchemaChartData


class PPTXTester:
    def __init__(self, test_file: str) -> None:
        self.test_file = test_file
        self.manager = PPTXManager(test_file)
        self.temp_dir = "test_temp"
        self._setup_environment()

        # Test counters
        self.total_tests = 0
        self.passed_tests = 0
        self.failed_tests = 0

    def _setup_environment(self) -> None:
        Path(self.temp_dir).mkdir(exist_ok=True)
        self.output_pptx = str(Path(self.temp_dir) / "output.pptx")
        self.output_pdf = str(Path(self.temp_dir) / "output.pdf")

    def _print_result(self, test_name: str, success: bool, error: str = "") -> None:
        self.total_tests += 1
        status = "PASSED" if success else "FAILED"
        self.passed_tests += 1 if success else 0
        self.failed_tests += 0 if success else 1

        print(f"[{status}] {test_name}")
        if not success and error:
            print(f"  └─ Error: {error}")

    def run_all_tests(self) -> None:
        print(f"Running PPTX Manager Tests on {self.test_file}\n" + "="*50)

        # Core functionality
        self.test_get_presentation_info()
        self.test_slide_analysis()

        # Slide operations
        self.test_slide_management()
        self.test_background_operations()

        # Content manipulation
        self.test_text_operations()
        self.test_image_operations()
        self.test_chart_operations()
        self.test_table_operations()

        # Export and cleanup
        self.test_save_functionality()

        # Final results
        print("\n" + "="*50)
        print(f"Total Tests: {self.total_tests}")
        print(f"Passed: {self.passed_tests} | Failed: {self.failed_tests}")
        print(f"Success Rate: {self.passed_tests/self.total_tests:.0%}")

    #region Test Methods
    def test_get_presentation_info(self) -> None:
        test_name = "Get Presentation Metadata"
        try:
            info = self.manager.get_presentation_info()
            assert info.slides_count > 0, "No slides found"
            assert info.dimensions.width > 0, "Invalid slide width"
            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_slide_analysis(self) -> None:
        test_name = "Slide Content Analysis"
        try:
            slide = self.manager.get_slide_details(1)
            assert len(slide.shapes) > 0, "No shapes in first slide"
            design = self.manager.analyze_slide_design(1)
            assert len(design.theme_colors) > 0, "No theme colors detected"
            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_slide_management(self) -> None:
        try:
            initial_count = len(self.manager.prs.slides)
            if initial_count == 0:
                self.manager.create_new_slide("Title Slide")
                initial_count += 1

            new_count = self.manager.duplicate_slide(1)
            self.manager.delete_slide(new_count)
            self._print_result("Slide Operations", True)
        except Exception as e:
            self._print_result("Slide Operations", False, str(e))

    def test_background_operations(self) -> None:
        test_name = "Background Modification"
        try:
            original_bg = self.manager.get_slide_details(1).background
            test_bg = Background(type="color", value="#FF0000", transparency=0.5)

            self.manager.set_slide_background(1, test_bg)
            updated_bg = self.manager.get_slide_details(1).background

            assert updated_bg.value == test_bg.value, "Background not updated"
            self.manager.set_slide_background(1, original_bg)

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_text_operations(self) -> None:
        test_name = "Text Manipulation"
        try:
            # Add text
            text_id = self.manager.add_text_block(
                1,
                "Test Text",
                Coordinates(x=1, y=1, width=3, height=1),
                FontStyle(color="#00FF00", bold=True)
            )

            # Modify text
            self.manager.edit_text_content(1, text_id, "Updated Text")
            slide = self.manager.get_slide_details(1)
            text_shapes = [s for s in slide.shapes if isinstance(s, TextShape)]
            assert any(t.content.text == "Updated Text" for t in text_shapes)

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_image_operations(self) -> None:
        test_name = "Image Handling"
        try:
            # Insert test image
            test_image = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwAEpwGkN600QAAAAABJRU5ErkJggg==")
            img_id = self.manager.insert_image(
                1,
                test_image,
                Coordinates(x=2, y=2, width=1, height=1)
            )

            # Replace image
            self.manager.replace_image(1, img_id, test_image)

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_chart_operations(self) -> None:
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

            # Modify data
            new_data = SchemaChartData(
                categories=["Q1", "Q2", "Q3", "Q4"],
                series=[{"name": "Sales", "values": [150, 250, 200, 350]}]
            )
            self.manager.modify_chart_data(1, chart_id, new_data)

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_table_operations(self) -> None:
        test_name = "Table Operations"
        try:
            # Create table
            table_id = self.manager.create_table(
                1, 3, 3,
                Coordinates(x=4, y=4, width=3, height=2)
            )

            # Edit cell
            cell = TableCell(row=0, col=0, content="Header", span_cols=3)
            self.manager.edit_table_cell(1, table_id, cell)

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))

    def test_save_functionality(self) -> None:
        test_name = "Presentation Save"
        try:
            self.manager.save(self.output_pptx)
            assert Path(self.output_pptx).exists(), "File not saved"
            assert os.path.getsize(self.output_pptx) > 1024, "Empty file saved"

            self._print_result(test_name, True)
        except Exception as e:
            self._print_result(test_name, False, str(e))
    #endregion

if __name__ == "__main__":
    test_file = "test_sources/test_dit.pptx"
    if not Path(test_file).exists():
        print(f"Test file not found: {test_file}")
        sys.exit(1)

    tester = PPTXTester(test_file)
    tester.run_all_tests()
