import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton, QVBoxLayout, QHBoxLayout

import csv_reader
import image_processor
import image_namer
import automation

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Create input fields and labels
        run_name_label = QLabel("Run Name:")
        self.run_name_input = QLineEdit()

        xy_count_label = QLabel("Number of XY Sequences:")
        self.xy_count_input = QLineEdit()

        naming_template_label = QLabel("Naming Template:")
        self.naming_template_input = QLineEdit()

        stitch_type_label = QLabel("Stitch Type:")
        self.stitch_type_combo = QComboBox()
        self.stitch_type_combo.addItems(["Full", "Load"])

        overlay_label = QLabel("Overlay Image:")
        self.overlay_checkbox = QCheckBox()

        submit_button = QPushButton("Submit")
        submit_button.clicked.connect(self.process_inputs)

        # Create layout
        layout = QVBoxLayout()
        layout.addWidget(run_name_label)
        layout.addWidget(self.run_name_input)
        layout.addWidget(xy_count_label)
        layout.addWidget(self.xy_count_input)
        layout.addWidget(naming_template_label)
        layout.addWidget(self.naming_template_input)
        layout.addWidget(stitch_type_label)
        layout.addWidget(self.stitch_type_combo)

        overlay_layout = QHBoxLayout()
        overlay_layout.addWidget(overlay_label)
        overlay_layout.addWidget(self.overlay_checkbox)
        layout.addLayout(overlay_layout)

        layout.addWidget(submit_button)

        self.setLayout(layout)
        self.setWindowTitle("BZ-X800 Analyzer")

    def process_inputs(self):
        run_name = self.run_name_input.text()
        xy_count = int(self.xy_count_input.text())
        naming_template = self.naming_template_input.text()
        stitch_type = self.stitch_type_combo.currentText()
        overlay = "Y" if self.overlay_checkbox.isChecked() else "N"

        placeholder_values = csv_reader.get_placeholder_values(xy_count, naming_template)
        main_window, stitch_button = automation.set_up_app()
        channel_orders_list = automation.define_channel_orders()

        for i in range(xy_count):
            xy_name = f"XY{i+1:02}"
            automation.process_xy_sequence(main_window, run_name, xy_name, stitch_button, stitch_type, overlay)
            image_processor.check_for_image(overlay)
            delay_time = image_processor.wait_for_viewer()
            image_namer.name_files(naming_template, placeholder_values, xy_name, delay_time, channel_orders_list)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())