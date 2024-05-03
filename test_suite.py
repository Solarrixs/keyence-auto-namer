import unittest
import os
from unittest.mock import patch, Mock
from config import Config
from keyence_analyzer import KeyenceAnalyzer
from wide_image_viewer import WideImageViewer
from main import Main

# The Keyence software must be opened for the test cases to run! pywinauto must be connected to the Keyence software to setup the Application object.

class TestKeyenceAutoNamer(unittest.TestCase):
    def setUp(self):
        self.config = Config("test_run", "F", "Y", "{key1}_{C}", "test_path", "image.png", "image2.png", 20)
        self.keyence_analyzer = KeyenceAnalyzer()
        self.wide_image_viewer = WideImageViewer()
        self.main = Main()

    def test_get_config(self):
        with patch('builtins.input', side_effect=["test_run", "F", "Y", "{key1}_{C}", "test_path"]):
            config = self.main.get_config()
            self.assertIsInstance(config, Config)
            self.assertEqual(config.run_name, "test_run")
            self.assertEqual(config.stitchtype, "F")
            self.assertEqual(config.overlay, "Y")
            self.assertEqual(config.naming_template, "{key1}_{C}")
            self.assertEqual(config.filepath, "test_path")
            self.assertEqual(config.IMAGE_PATH, os.path.join(os.path.dirname(__file__), 'image.png'))
            self.assertEqual(config.IMAGE_PATH_2, os.path.join(os.path.dirname(__file__), 'image2.png'))
            self.assertEqual(config.MAX_DELAY_TIME, 20)

    def test_get_config_empty_inputs(self):
        with patch('builtins.input', side_effect=["", "", "", "", ""]):
            with self.assertRaises(ValueError):
                self.main.get_config()

    def test_get_xy_sequence_range(self):
        with patch('builtins.input', side_effect=["1", "5"]):
            start_child, end_child = self.main.get_xy_sequence_range("test_run")
            self.assertEqual(start_child, 1)
            self.assertEqual(end_child, 5)

    def test_get_xy_sequence_range_invalid_input(self):
        with patch('builtins.input', side_effect=["0", "5"]):
            with self.assertRaises(ValueError):
                self.main.get_xy_sequence_range("test_run")

    def test_get_placeholder_values(self):
        with patch('builtins.input', side_effect=["value1", "value2"]):
            placeholder_values = self.main.get_placeholder_values("{key1}_{key2}_{C}", 1, 2)
            expected_values = {
                "XY01": {"key1": "value1", "key2": "value2"},
                "XY02": {"key1": "value1", "key2": "value2"}
            }
            self.assertEqual(placeholder_values, expected_values)

    def test_get_placeholder_values_empty_template(self):
        placeholder_values = self.main.get_placeholder_values("", 1, 2)
        self.assertEqual(placeholder_values, {})

    def test_define_channel_orders(self):
        with patch('builtins.input', side_effect=["3", "C1", "C2", "C3"]):
            channel_orders_list = self.main.define_channel_orders()
            self.assertEqual(channel_orders_list, ["C3", "C2", "C1"])

    def test_define_channel_orders_invalid_count(self):
        with patch('builtins.input', side_effect=["0"]):
            with self.assertRaises(ValueError):
                self.main.define_channel_orders()

    def test_select_stitch_type_full_focus(self):
        with patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.select_stitch_type("F")
            mock_press.assert_any_call('f')
            mock_press.assert_any_call('enter')

    def test_select_stitch_type_load(self):
        with patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.select_stitch_type("L")
            mock_press.assert_called_once_with('l')

    def test_select_stitch_type_invalid(self):
        with patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.select_stitch_type("X")
            mock_press.assert_not_called()

    def test_check_for_image_exists(self):
        with patch('os.path.exists', return_value=True), \
             patch('pyautogui.locateOnScreen', return_value=(0, 0, 100, 100)), \
             patch('pyautogui.click') as mock_click:
            self.keyence_analyzer.check_for_image("image.png")
            mock_click.assert_called_once()

    def test_check_for_image_not_found(self):
        with patch('os.path.exists', return_value=True), \
             patch('pyautogui.locateOnScreen', return_value=None), \
             patch('time.sleep', side_effect=Exception("Image not found")):
            with self.assertRaises(Exception):
                self.keyence_analyzer.check_for_image("image.png")

    def test_start_stitching_with_overlay(self):
        with patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.start_stitching("Y")
            mock_press.assert_any_call('tab', presses=6)
            mock_press.assert_any_call('right')
            mock_press.assert_any_call('tab', presses=3)
            mock_press.assert_any_call('enter')

    def test_start_stitching_without_overlay(self):
        with patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.start_stitching("N")
            mock_press.assert_any_call('tab', presses=6)
            mock_press.assert_any_call('right')
            mock_press.assert_any_call('tab', presses=2)
            mock_press.assert_any_call('enter')

    def test_disable_caps_lock_enabled(self):
        with patch('ctypes.windll.user32.GetKeyState', return_value=1), \
             patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.disable_caps_lock()
            mock_press.assert_called_once_with('capslock')

    def test_disable_caps_lock_disabled(self):
        with patch('ctypes.windll.user32.GetKeyState', return_value=0), \
             patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.disable_caps_lock()
            mock_press.assert_not_called()

    def test_close_stitch_image(self):
        with patch('time.sleep'), \
             patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.close_stitch_image(1)
            mock_press.assert_any_call('tab', presses=2)
            mock_press.assert_any_call('enter')

    def test_wait_for_viewer(self):
        with patch('pywinauto.Desktop.windows', return_value=[]), \
             patch('time.sleep', side_effect=Exception("Viewer not found")):
            with self.assertRaises(Exception):
                self.wide_image_viewer.wait_for_viewer(5)

    def test_name_files(self):
        with patch('pywinauto.Desktop.windows', return_value=[]), \
             patch('pyautogui.press'), \
             patch('pyautogui.write'), \
             patch('time.sleep'), \
             patch('WideImageViewer.click_file_button'), \
             patch('WideImageViewer.export_in_original_scale'), \
             patch('WideImageViewer.close_image'):
            self.wide_image_viewer.name_files("{key1}_{C}", {"XY01": {"key1": "value1"}}, "XY01", 1, "", ["C1"])

    def test_click_file_button(self):
        with patch('pywinauto.Application.connect'), \
             patch('pywinauto.Application.window'), \
             patch('pywinauto.Application.window.child_window'), \
             patch('pywinauto.Application.window.child_window.click_input'):
            self.wide_image_viewer.click_file_button(None)

    def test_export_in_original_scale(self):
        with patch('pyautogui.press') as mock_press:
            self.wide_image_viewer.export_in_original_scale()
            mock_press.assert_any_call('tab', presses=4)
            mock_press.assert_any_call('enter')
            mock_press.assert_any_call('tab', presses=1)
            mock_press.assert_any_call('enter')

    def test_close_image(self):
        with patch('time.sleep'), \
             patch('pyautogui.hotkey'), \
             patch('pyautogui.press'):
            self.wide_image_viewer.close_image(1, "C1")
            
    def test_get_config_invalid_stitchtype(self):
        with patch('builtins.input', side_effect=["test_run", "X", "Y", "{key1}_{C}", "test_path"]):
            with self.assertRaises(ValueError):
                self.main.get_config()

    def test_get_config_invalid_overlay(self):
        with patch('builtins.input', side_effect=["test_run", "F", "X", "{key1}_{C}", "test_path"]):
            with self.assertRaises(ValueError):
                self.main.get_config()

    def test_get_config_invalid_template(self):
        with patch('builtins.input', side_effect=["test_run", "F", "Y", "invalid_template", "test_path"]):
            with self.assertRaises(ValueError):
                self.main.get_config()

    def test_get_xy_sequence_range_negative_values(self):
        with patch('builtins.input', side_effect=["-1", "-5"]):
            with self.assertRaises(ValueError):
                self.main.get_xy_sequence_range("test_run")

    def test_get_xy_sequence_range_start_greater_than_end(self):
        with patch('builtins.input', side_effect=["5", "1"]):
            with self.assertRaises(ValueError):
                self.main.get_xy_sequence_range("test_run")

    def test_get_xy_sequence_range_non_numeric_input(self):
        with patch('builtins.input', side_effect=["abc", "xyz"]):
            with self.assertRaises(ValueError):
                self.main.get_xy_sequence_range("test_run")

    def test_get_placeholder_values_special_characters(self):
        with patch('builtins.input', side_effect=["value!@#", "value$%^"]):
            placeholder_values = self.main.get_placeholder_values("{key1}_{key2}_{C}", 1, 2)
            expected_values = {
                "XY01": {"key1": "value!@#", "key2": "value$%^"},
                "XY02": {"key1": "value!@#", "key2": "value$%^"}
            }
            self.assertEqual(placeholder_values, expected_values)

    def test_get_placeholder_values_empty_input(self):
        with patch('builtins.input', side_effect=["", ""]):
            placeholder_values = self.main.get_placeholder_values("{key1}_{key2}_{C}", 1, 2)
            expected_values = {
                "XY01": {"key1": "", "key2": ""},
                "XY02": {"key1": "", "key2": ""}
            }
            self.assertEqual(placeholder_values, expected_values)

    def test_define_channel_orders_empty_input(self):
        with patch('builtins.input', side_effect=["3", "", "", ""]):
            channel_orders_list = self.main.define_channel_orders()
            self.assertEqual(channel_orders_list, ["", "", ""])

    def test_define_channel_orders_special_characters(self):
        with patch('builtins.input', side_effect=["3", "C!@#", "C$%^", "C&*("]):
            channel_orders_list = self.main.define_channel_orders()
            self.assertEqual(channel_orders_list, ["C&*(", "C$%^", "C!@#"])

    def test_check_for_image_file_not_found(self):
        with patch('os.path.exists', return_value=False):
            with self.assertRaises(AssertionError):
                self.keyence_analyzer.check_for_image("nonexistent_image.png")

    def test_check_for_image_timeout(self):
        with patch('os.path.exists', return_value=True), \
             patch('pyautogui.locateOnScreen', return_value=None), \
             patch('time.sleep'), \
             patch('time.time', side_effect=[0, 601]):
            with self.assertRaises(SystemExit):
                self.keyence_analyzer.check_for_image("image.png")

    def test_close_stitch_image_zero_delay(self):
        with patch('time.sleep'), \
             patch('pyautogui.press') as mock_press:
            self.keyence_analyzer.close_stitch_image(0)
            mock_press.assert_any_call('tab', presses=2)
            mock_press.assert_any_call('enter')

    def test_wait_for_viewer_immediate_found(self):
        with patch('pywinauto.Desktop.windows', return_value=[Mock(window_text=self.wide_image_viewer.title)]):
            delay_time = self.wide_image_viewer.wait_for_viewer(5)
            self.assertAlmostEqual(delay_time, 0, places=1)

    def test_wait_for_viewer_timeout(self):
        with patch('pywinauto.Desktop.windows', return_value=[]), \
             patch('time.sleep'), \
             patch('time.time', side_effect=[0, 21]):
            delay_time = self.wide_image_viewer.wait_for_viewer(20)
            self.assertEqual(delay_time, 20)

    def test_name_files_invalid_template(self):
        with self.assertRaises(KeyError):
            self.wide_image_viewer.name_files("invalid_template", {"XY01": {}}, "XY01", 1, "", ["C1"])

    def test_click_file_button_window_not_found(self):
        with patch('pywinauto.Application.connect'), patch('pywinauto.Application.window', side_effect=Exception("Window not found")):
            with self.assertRaises(Exception):
                self.wide_image_viewer.click_file_button(None)

if __name__ == '__main__':
    unittest.main()