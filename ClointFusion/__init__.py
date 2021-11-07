# ---------  Libraries Imports | Current Count : 15
try:
    import platform
    import struct
    os_name = str(platform.system()).lower()
    bit_size = struct.calcsize("P") * 8
    if bit_size == 32:
        print("We don't support 32-bit Architecture")
        print("Please download 64-bit Architecture from the link down below. Press (ctrl + click) on link to download.\n")
        if os_name == "windows":
            print("https://www.python.org/ftp/python/3.9.7/python-3.9.7-amd64.exe")
        if os_name == "darwin":
            print("https://www.python.org/ftp/python/3.9.7/python-3.9.7-macos11.pkg")
        sys.exit(0)
    else:
        from ClointFusion.cce import pd
        try:
            from ClointFusion.cce import pg
        except:
            pass

        from ClointFusion.cce import clipboard
        from ClointFusion.cce import re
        from ClointFusion.cce import op
        from ClointFusion.cce import os
        from ClointFusion.cce import time
        from ClointFusion.cce import shutil
        from ClointFusion.cce import sys
        from ClointFusion.cce import datetime
        from ClointFusion.cce import subprocess
        from ClointFusion.cce import traceback
        from ClointFusion.cce import logging
        from ClointFusion.cce import user_name
        from ClointFusion.cce import user_email
        from ClointFusion.cce import webbrowser

        # ---------  Variables Imports | Current Count : 8
        from ClointFusion.cce import batch_file_path
        from ClointFusion.cce import output_folder_path
        from ClointFusion.cce import config_folder_path
        from ClointFusion.cce import img_folder_path
        from ClointFusion.cce import error_screen_shots_path
        from ClointFusion.cce import cf_icon_file_path
        from ClointFusion.cce import cf_logo_file_path
        from ClointFusion.cce import os_name

        # ---------  GUI Functions | Current Count : 6
        from ClointFusion.cce import gui_get_consent_from_user
        from ClointFusion.cce import gui_get_dropdownlist_values_from_user
        from ClointFusion.cce import gui_get_excel_sheet_header_from_user
        from ClointFusion.cce import gui_get_folder_path_from_user
        from ClointFusion.cce import gui_get_any_input_from_user
        from ClointFusion.cce import gui_get_any_file_from_user


        # ---------  Message  Functions | Current Count : 4
        from ClointFusion.cce import message_counter_down_timer
        from ClointFusion.cce import message_pop_up
        from ClointFusion.cce import message_flash
        from ClointFusion.cce import message_toast


        # ---------  Mouse Functions | Current Count : 4
        from ClointFusion.cce import mouse_click
        from ClointFusion.cce import mouse_move
        from ClointFusion.cce import mouse_drag_from_to
        from ClointFusion.cce import mouse_search_snip_return_coordinates_x_y


        # ---------  Keyboard Functions | Current Count : 3
        from ClointFusion.cce import key_press
        from ClointFusion.cce import key_write_enter
        from ClointFusion.cce import key_hit_enter


        # ---------  Browser Functions | Current Count : 11
        # from ClointFusion.cce import browser
        from ClointFusion.cce import browser_activate
        from ClointFusion.cce import browser_navigate_h
        from ClointFusion.cce import browser_write_h
        from ClointFusion.cce import browser_mouse_click_h
        from ClointFusion.cce import browser_locate_element_h
        from ClointFusion.cce import browser_wait_until_h
        from ClointFusion.cce import browser_refresh_page_h
        from ClointFusion.cce import browser_hit_enter_h
        from ClointFusion.cce import browser_key_press_h
        from ClointFusion.cce import browser_mouse_hover_h
        from ClointFusion.cce import browser_quit_h
        from ClointFusion.cce import browser_set_waiting_time


        # ---------  Folder Functions | Current Count : 8
        from ClointFusion.cce import folder_read_text_file
        from ClointFusion.cce import folder_write_text_file
        from ClointFusion.cce import folder_create
        from ClointFusion.cce import folder_create_text_file
        from ClointFusion.cce import folder_get_all_filenames_as_list
        from ClointFusion.cce import folder_delete_all_files
        from ClointFusion.cce import file_rename
        from ClointFusion.cce import file_get_json_details


        # ---------  Window Operations Functions | Current Count : 6
        from ClointFusion.cce import window_show_desktop
        from ClointFusion.cce import window_get_all_opened_titles_windows
        from ClointFusion.cce import window_activate_and_maximize_windows
        from ClointFusion.cce import window_minimize_windows
        from ClointFusion.cce import window_close_windows
        from ClointFusion.cce import launch_any_exe_bat_application


        # ---------  String Functions | Current Count : 3
        from ClointFusion.cce import string_extract_only_alphabets
        from ClointFusion.cce import string_extract_only_numbers
        from ClointFusion.cce import string_remove_special_characters


        # ---------  Excel Functions | Current Count : 28
        from ClointFusion.cce import excel_get_row_column_count
        from ClointFusion.cce import excel_copy_range_from_sheet
        from ClointFusion.cce import excel_copy_paste_range_from_to_sheet
        from ClointFusion.cce import excel_split_by_column
        from ClointFusion.cce import excel_split_the_file_on_row_count
        from ClointFusion.cce import excel_merge_all_files
        from ClointFusion.cce import excel_drop_columns
        from ClointFusion.cce import excel_sort_columns
        from ClointFusion.cce import excel_clear_sheet
        from ClointFusion.cce import excel_set_single_cell
        from ClointFusion.cce import excel_get_single_cell
        from ClointFusion.cce import excel_remove_duplicates
        from ClointFusion.cce import excel_vlook_up
        from ClointFusion.cce import excel_change_corrupt_xls_to_xlsx
        from ClointFusion.cce import excel_convert_xls_to_xlsx
        from ClointFusion.cce import excel_apply_format_as_table
        from ClointFusion.cce import excel_split_on_user_defined_conditions
        from ClointFusion.cce import excel_convert_to_image
        from ClointFusion.cce import excel_create_excel_file_in_given_folder
        from ClointFusion.cce import excel_if_value_exists
        from ClointFusion.cce import excel_create_file
        from ClointFusion.cce import excel_to_colored_html
        from ClointFusion.cce import excel_get_all_sheet_names
        from ClointFusion.cce import excel_get_all_header_columns
        from ClointFusion.cce import excel_describe_data
        from ClointFusion.cce import excel_sub_routines
        from ClointFusion.cce import convert_csv_to_excel
        from ClointFusion.cce import isNaN


        # --------- Windows Objects Functions | Current Count : 5
        from ClointFusion.cce import win_obj_open_app
        from ClointFusion.cce import win_obj_get_all_objects
        from ClointFusion.cce import win_obj_mouse_click
        from ClointFusion.cce import win_obj_key_press
        from ClointFusion.cce import win_obj_get_text


        # --------- Screenscraping Functions | Current Count : 5
        from ClointFusion.cce import scrape_save_contents_to_notepad
        from ClointFusion.cce import scrape_get_contents_by_search_copy_paste
        from ClointFusion.cce import screen_clear_search
        from ClointFusion.cce import search_highlight_tab_enter_open
        from ClointFusion.cce import find_text_on_screen


        # --------- Schedule Functions | Current Count : 2
        from ClointFusion.cce import schedule_create_task_windows
        from ClointFusion.cce import schedule_delete_task_windows


        # --------- Email Functions | Current Count : 1
        from ClointFusion.cce import email_send_via_desktop_outlook


        # --------- Utility Functions | Current Count : 12
        from ClointFusion.cce import find
        from ClointFusion.cce import pause_program
        from ClointFusion.cce import show_emoji
        from ClointFusion.cce import create_batch_file
        from ClointFusion.cce import dismantle_code
        from ClointFusion.cce import download_this_file
        from ClointFusion.cce import clear_screen
        from ClointFusion.cce import print_with_magic_color
        from ClointFusion.cce import ocr_now
        from ClointFusion.cce import string_regex

        # --------- Self-Test and ClointFusion Related Functions | Current Count : 3
        from ClointFusion.cce import update_log_excel_file
        from ClointFusion.cce import ON_semi_automatic_mode
        from ClointFusion.cce import OFF_semi_automatic_mode

        # Voice Interface
        from ClointFusion.cce import text_to_speech
        from ClointFusion.cce import speech_to_text
except:
    pass
