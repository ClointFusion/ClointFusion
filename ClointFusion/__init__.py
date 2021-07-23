from ClointFusion import loader    

# ---------  Libraries Imports | Current Count : 13
from loader import pd
try:
    from loader import pg
except:
    pass
from loader import clipboard
from loader import re
from loader import op
from loader import os
from loader import time
from loader import shutil
from loader import sys
from loader import datetime
from loader import subprocess
from loader import traceback
from loader import logging

# ---------  Variables Imports | Current Count : 8
from loader import batch_file_path
from loader import output_folder_path
from loader import config_folder_path
from loader import img_folder_path
from loader import error_screen_shots_path
from loader import cf_icon_file_path
from loader import cf_logo_file_path
from loader import os_name


# ---------  GUI Functions | Current Count : 6
from loader import gui_get_consent_from_user
from loader import gui_get_dropdownlist_values_from_user
from loader import gui_get_excel_sheet_header_from_user
from loader import gui_get_folder_path_from_user
from loader import gui_get_any_input_from_user
from loader import gui_get_any_file_from_user


# ---------  Message  Functions | Current Count : 4
from loader import message_counter_down_timer
from loader import message_pop_up
from loader import message_flash
from loader import message_toast


# ---------  Mouse Functions | Current Count : 5
from loader import mouse_click
from loader import mouse_move
from loader import mouse_get_color_by_position
from loader import mouse_drag_from_to
from loader import mouse_search_snip_return_coordinates_x_y


# ---------  Keyboard Functions | Current Count : 3
from loader import key_press
from loader import key_write_enter
from loader import key_hit_enter


# ---------  Browser Functions | Current Count : 11
from loader import browser_activate
from loader import browser_navigate_h
from loader import browser_write_h
from loader import browser_mouse_click_h
from loader import browser_locate_element_h
from loader import browser_wait_until_h
from loader import browser_refresh_page_h
from loader import browser_hit_enter_h
from loader import browser_key_press_h
from loader import browser_mouse_hover_h
from loader import browser_quit_h


# ---------  Folder Functions | Current Count : 8
from loader import folder_read_text_file
from loader import folder_write_text_file
from loader import folder_create
from loader import folder_create_text_file
from loader import folder_get_all_filenames_as_list
from loader import folder_delete_all_files
from loader import file_rename
from loader import file_get_json_details


# ---------  Window Operations Functions | Current Count : 6
from loader import window_show_desktop
from loader import window_get_all_opened_titles_windows
from loader import window_activate_and_maximize_windows
from loader import window_minimize_windows
from loader import window_close_windows
from loader import launch_any_exe_bat_application


# ---------  String Functions | Current Count : 3
from loader import string_extract_only_alphabets
from loader import string_extract_only_numbers
from loader import string_remove_special_characters


# ---------  Excel Functions | Current Count : 29
from loader import excel_get_row_column_count
from loader import excel_copy_range_from_sheet
from loader import excel_copy_paste_range_from_to_sheet
from loader import excel_split_by_column
from loader import excel_split_the_file_on_row_count
from loader import excel_merge_all_files
from loader import excel_drop_columns
from loader import excel_sort_columns
from loader import excel_clear_sheet
from loader import excel_set_single_cell
from loader import excel_get_single_cell
from loader import excel_remove_duplicates
from loader import excel_vlook_up
from loader import excel_change_corrupt_xls_to_xlsx
from loader import excel_convert_xls_to_xlsx
from loader import excel_apply_template_format_save_to_new
from loader import excel_apply_format_as_table
from loader import excel_split_on_user_defined_conditions
from loader import excel_convert_to_image
from loader import excel_create_excel_file_in_given_folder
from loader import excel_if_value_exists
from loader import excel_create_file
from loader import excel_to_colored_html
from loader import excel_get_all_sheet_names
from loader import excel_get_all_header_columns
from loader import excel_describe_data
from loader import excel_sub_routines
from loader import convert_csv_to_excel
from loader import isNaN


# --------- Windows Objects Functions | Current Count : 5
from loader import win_obj_open_app
from loader import win_obj_get_all_objects
from loader import win_obj_mouse_click
from loader import win_obj_key_press
from loader import win_obj_get_text


# --------- Screenscraping Functions | Current Count : 5
from loader import scrape_save_contents_to_notepad
from loader import scrape_get_contents_by_search_copy_paste
from loader import screen_clear_search
from loader import search_highlight_tab_enter_open
from loader import find_text_on_screen


# --------- Schedule Functions | Current Count : 2
from loader import schedule_create_task_windows
from loader import schedule_delete_task_windows


# --------- Email Functions | Current Count : 1
from loader import email_send_via_desktop_outlook


# --------- Utility Functions | Current Count : 9

from loader import find
from loader import pause_program
from loader import show_emoji
from loader import create_batch_file
from loader import dismantle_code
from loader import compute_hash
from loader import date_convert_to_US_format
from loader import download_this_file
from loader import get_image_from_base64
from loader import clear_screen


# --------- Self-Test and ClointFusion Related Functions | Current Count : 4
from loader import take_error_screenshot
from loader import update_log_excel_file
from loader import ON_semi_automatic_mode
from loader import OFF_semi_automatic_mode