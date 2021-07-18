# ---------  Libraries Imports | Current Count : 13
from ClointFusion.ClointFusion import pd
from ClointFusion.ClointFusion import pg
from ClointFusion.ClointFusion import clipboard
from ClointFusion.ClointFusion import re
from ClointFusion.ClointFusion import op
from ClointFusion.ClointFusion import os
from ClointFusion.ClointFusion import time
from ClointFusion.ClointFusion import shutil
from ClointFusion.ClointFusion import sys
from ClointFusion.ClointFusion import datetime
from ClointFusion.ClointFusion import subprocess
from ClointFusion.ClointFusion import traceback
from ClointFusion.ClointFusion import logging


# ---------  Method Imports | Current Count : 2
from ClointFusion.ClointFusion import background
from ClointFusion.ClointFusion import timeit


# ---------  Variables Imports | Current Count : 8
from ClointFusion.ClointFusion import batch_file_path
from ClointFusion.ClointFusion import output_folder_path
from ClointFusion.ClointFusion import config_folder_path
from ClointFusion.ClointFusion import img_folder_path
from ClointFusion.ClointFusion import error_screen_shots_path
from ClointFusion.ClointFusion import cf_icon_file_path
from ClointFusion.ClointFusion import cf_logo_file_path
from ClointFusion.ClointFusion import os_name


# ---------  GUI Functions | Current Count : 6
from ClointFusion.ClointFusion import gui_get_consent_from_user
from ClointFusion.ClointFusion import gui_get_dropdownlist_values_from_user
from ClointFusion.ClointFusion import gui_get_excel_sheet_header_from_user
from ClointFusion.ClointFusion import gui_get_folder_path_from_user
from ClointFusion.ClointFusion import gui_get_any_input_from_user
from ClointFusion.ClointFusion import gui_get_any_file_from_user


# ---------  Message  Functions | Current Count : 4
from ClointFusion.ClointFusion import message_counter_down_timer
from ClointFusion.ClointFusion import message_pop_up
from ClointFusion.ClointFusion import message_flash
from ClointFusion.ClointFusion import message_toast


# ---------  Mouse Functions | Current Count : 5
from ClointFusion.ClointFusion import mouse_click
from ClointFusion.ClointFusion import mouse_move
from ClointFusion.ClointFusion import mouse_get_color_by_position
from ClointFusion.ClointFusion import mouse_drag_from_to
from ClointFusion.ClointFusion import mouse_search_snip_return_coordinates_x_y


# ---------  Keyboard Functions | Current Count : 3
from ClointFusion.ClointFusion import keyboard_hotkey_press
from ClointFusion.ClointFusion import keyboard_write_text
from ClointFusion.ClointFusion import keyboard_hit_enter


# ---------  Browser Functions | Current Count : 11
from ClointFusion.ClointFusion import browser_activate
from ClointFusion.ClointFusion import browser_navigate_h
from ClointFusion.ClointFusion import browser_write_h
from ClointFusion.ClointFusion import browser_mouse_click_h
from ClointFusion.ClointFusion import browser_locate_element_h
from ClointFusion.ClointFusion import browser_wait_until_h
from ClointFusion.ClointFusion import browser_refresh_page_h
from ClointFusion.ClointFusion import browser_hit_enter_h
from ClointFusion.ClointFusion import browser_key_press_h
from ClointFusion.ClointFusion import browser_mouse_hover_h
from ClointFusion.ClointFusion import browser_quit_h


# ---------  Folder Functions | Current Count : 8
from ClointFusion.ClointFusion import folder_read_text_file
from ClointFusion.ClointFusion import folder_write_text_file
from ClointFusion.ClointFusion import folder_create
from ClointFusion.ClointFusion import folder_create_text_file
from ClointFusion.ClointFusion import folder_get_all_filenames_as_list
from ClointFusion.ClointFusion import folder_delete_all_files
from ClointFusion.ClointFusion import file_rename
from ClointFusion.ClointFusion import file_get_json_details


# ---------  Window Operations Functions | Current Count : 6
from ClointFusion.ClointFusion import window_show_desktop
from ClointFusion.ClointFusion import window_get_all_opened_titles_windows
from ClointFusion.ClointFusion import window_activate_and_maximize_windows
from ClointFusion.ClointFusion import window_minimize_windows
from ClointFusion.ClointFusion import window_close_windows
from ClointFusion.ClointFusion import launch_any_exe_bat_application


# ---------  String Functions | Current Count : 3
from ClointFusion.ClointFusion import string_extract_only_alphabets
from ClointFusion.ClointFusion import string_extract_only_numbers
from ClointFusion.ClointFusion import string_remove_special_characters


# ---------  Excel Functions | Current Count : 29
from ClointFusion.ClointFusion import excel_get_row_column_count
from ClointFusion.ClointFusion import excel_copy_range_from_sheet
from ClointFusion.ClointFusion import excel_copy_paste_range_from_to_sheet
from ClointFusion.ClointFusion import excel_split_by_column
from ClointFusion.ClointFusion import excel_split_the_file_on_row_count
from ClointFusion.ClointFusion import excel_merge_all_files
from ClointFusion.ClointFusion import excel_drop_columns
from ClointFusion.ClointFusion import excel_sort_columns
from ClointFusion.ClointFusion import excel_clear_sheet
from ClointFusion.ClointFusion import excel_set_single_cell
from ClointFusion.ClointFusion import excel_get_single_cell
from ClointFusion.ClointFusion import excel_remove_duplicates
from ClointFusion.ClointFusion import excel_vlook_up
from ClointFusion.ClointFusion import excel_change_corrupt_xls_to_xlsx
from ClointFusion.ClointFusion import excel_convert_xls_to_xlsx
from ClointFusion.ClointFusion import excel_apply_template_format_save_to_new
from ClointFusion.ClointFusion import excel_apply_format_as_table
from ClointFusion.ClointFusion import excel_split_on_user_defined_conditions
from ClointFusion.ClointFusion import excel_convert_to_image
from ClointFusion.ClointFusion import excel_create_excel_file_in_given_folder
from ClointFusion.ClointFusion import excel_if_value_exists
from ClointFusion.ClointFusion import excel_create_file
from ClointFusion.ClointFusion import excel_to_colored_html
from ClointFusion.ClointFusion import excel_get_all_sheet_names
from ClointFusion.ClointFusion import excel_get_all_header_columns
from ClointFusion.ClointFusion import excel_describe_data
from ClointFusion.ClointFusion import excel_sub_routines
from ClointFusion.ClointFusion import convert_csv_to_excel
from ClointFusion.ClointFusion import isNaN


# --------- Windows Objects Functions | Current Count : 5
from ClointFusion.ClointFusion import win_obj_open_app
from ClointFusion.ClointFusion import win_obj_get_all_objects
from ClointFusion.ClointFusion import win_obj_mouse_click
from ClointFusion.ClointFusion import win_obj_key_press
from ClointFusion.ClointFusion import win_obj_get_text


# --------- Screenscraping Functions | Current Count : 5
from ClointFusion.ClointFusion import scrape_save_contents_to_notepad
from ClointFusion.ClointFusion import scrape_get_contents_by_search_copy_paste
from ClointFusion.ClointFusion import screen_clear_search
from ClointFusion.ClointFusion import search_highlight_tab_enter_open
from ClointFusion.ClointFusion import find_text_on_screen


# --------- Schedule Functions | Current Count : 2
from ClointFusion.ClointFusion import schedule_create_task_windows
from ClointFusion.ClointFusion import schedule_delete_task_windows


# --------- Email Functions | Current Count : 1
from ClointFusion.ClointFusion import email_send_via_desktop_outlook


# --------- Utility Functions | Current Count : 9
from ClointFusion.ClointFusion import find
from ClointFusion.ClointFusion import pause_program
from ClointFusion.ClointFusion import show_emoji
from ClointFusion.ClointFusion import create_batch_file
from ClointFusion.ClointFusion import dismantle_code
from ClointFusion.ClointFusion import compute_hash
from ClointFusion.ClointFusion import date_convert_to_US_format
from ClointFusion.ClointFusion import download_this_file
from ClointFusion.ClointFusion import get_image_from_base64


# --------- Self-Test and ClointFusion Related Functions | Current Count : 4
from ClointFusion.ClointFusion import take_error_screenshot
from ClointFusion.ClointFusion import update_log_excel_file
from ClointFusion.ClointFusion import ON_semi_automatic_mode
from ClointFusion.ClointFusion import OFF_semi_automatic_mode