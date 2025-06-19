*-------------------------------------------------------------------------*
*   Confidential Property of Stryker
*   All Rights Reserved
***************************************************************************
* Program name    : Report ZFTSR_INVLOC_TRANSFER                          *
* Company         : Stryker Project Accelerate                            *
* Author          : VSINGH                                                *
* Date            : Nov 26, 2018                                          *
* Title           : Inventory Location Transfer Report                    *
* FD #            : GLOBSAP_FTS.SCTASK1280071_Inventory Location Transfer *
***************************************************************************
*DESCRIPTION  :  The purpose of this report is Mass inventory transfer    *
*                for customer consignment fill up, pick up, initial       *
*                stock transfer and reversal.                             *
*                                                                         *
***************************************************************************
*         H I S T O R Y   O F   R E V I S I O N S                         *
***************************************************************************
*DATE        AUTHOR          DESCRIPTION OF CHANGE          Request #     *
***************************************************************************
*11/26/2018  VSINGH           Initial                       ECDK911792    *
*02/13/2023  VMANEM           R6 ENDO FTS.EXT.069           EC2K902748    *
*11/22/2023  VMANEM           R6-ENDO/BUG 669327            EC1K916745    *
* Display error message, when the posting is for customer consignment and *
* the file has storage location value                                     *
*04/22/2024  SDE2            R7-ORTHO FTS.EXT.286           EC2K904625    *
* Replace the Movement Type Radiobuttons with Dropdown,                   *
* Dynamic download template and additional checks before posting          *
* Include 551W and 552W movementy types                                   *
*02/07/2025  SASHOK       R7 - Fix for WebGUI file operation EC1K923128   *
***************************************************************************
*&------------------------------------------------------------------------*
*& Include          ZFTSR_INVLOC_TRANSFER_F03
*&------------------------------------------------------------------------*
* Initialization Event
INITIALIZATION.


* Method to perform authority check on Transaction
  CALL METHOD zcl_gbl_utility=>z_authority_check
    EXPORTING
      i_tcode = sy-tcode.

  sscrfields-functxt_01 = TEXT-001. " Download Template

* Creating object to call methods
  DATA(go_invloc_transfer) = NEW lcl_invloc_transfer( p_file ).

AT SELECTION-SCREEN.

  CONSTANTS : gc_range1  TYPE string VALUE zif_gbl_constants=>gc_range1,
              gc_range2  TYPE string VALUE zif_gbl_constants=>gc_range2,
              gc_headcol TYPE i VALUE zif_gbl_constants=>gc_headcol,
              gc_linecol TYPE i VALUE zif_gbl_constants=>gc_linecol.

  DATA(gv_file_upload) = NEW zcl_excel_file_upload( ).

**** Begin of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
  DATA : lv_file_gui TYPE rlgrap-filename. "For GUI
  DATA: lv_filepath TYPE string. "For WebGUI
  DATA : lv_win_title   TYPE string,
         lv_user_action TYPE i,
         lv_ctr_temp    TYPE string,
         lv_filename    TYPE string,
         lv_path        TYPE string,
         lv_full_path   TYPE string,
         lv_filter      TYPE string.
  CONSTANTS: lc_file_name_length TYPE i VALUE 255.
  DATA: lv_filename1 TYPE c LENGTH lc_file_name_length.

**** End of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
  CASE           sscrfields-ucomm.
    WHEN : zif_gbl_constants=>gc_fc01.

* Begin of Changes: R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
      IF p_mtyp IS INITIAL.
        MESSAGE e041.                                " Please Select Movement Type
      ELSE.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
*        go_invloc_transfer->z_create_template_header_table( ).                   "-- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

        " Introducing the new method for dynamic template creation
        go_invloc_transfer->z_create_dyn_template_hdr_tab( ).                     "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

*        DATA(lv_str) = gc_alphabet+gvi(1).
*        DATA(lv_string) = 'A1:' && lv_str.
        IF cl_gui_frontend_services=>www_active = abap_false. " ++ INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
          lv_file_gui = p_mtyp && '_' && TEXT-001.  " ++ INS R7-ORTHO/SASHOK/EC1K923128
          gv_file_upload->download_excel_file(
                           i_header_data  = go_invloc_transfer->lt_itab
                           i_filename     = lv_file_gui "p_file
*                         i_header_range = gc_range1                              "-- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
*                         i_item_range   = gc_range2                              "-- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
                           i_header_range =  'A1:' && gc_alphabet+gv_dynl(1) && '1' "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
                           i_header_color = gc_headcol ).
*                         i_item_color   = gc_linecol ).                          "-- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
**** Begin of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
        ELSE.

*& File Download in case of WEB GUI
          lv_filter    = 'Excel(*.xls )'(080).
          lv_win_title = TEXT-001 ."'Download Template'.
          lv_ctr_temp  = p_mtyp && '_' && TEXT-001.

          CALL METHOD gv_file_upload->z_download_save_dialog
            EXPORTING
              i_window_title = lv_win_title    " Window Title
              i_def_filename = lv_ctr_temp     " Template
              i_file_filter  = lv_filter       " File Type Filter Table
            CHANGING
              c_filename     = lv_filename     " File Name to Save
              c_path         = lv_path         " Path to File
              c_fullpath     = lv_full_path    " Path + File Name
              c_useraction   = lv_user_action. " User Action (C Class Const ACTION_OK, ACTION_OVERWRITE etc)

          IF lv_user_action NE cl_gui_frontend_services=>action_cancel.

            CALL METHOD gv_file_upload->z_download_excel_file_webgui
              EXPORTING
                i_filename    = lv_filename
              CHANGING
                c_header_data = go_invloc_transfer->lt_itab.
          ENDIF.
        ENDIF.
**** End of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
      ENDIF.                                                                      "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

  ENDCASE.

START-OF-SELECTION.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
  " Screen input check for Movement Type
  IF p_mtyp IS INITIAL.
    MESSAGE i041.                                " Please Select Movement Type
    LEAVE LIST-PROCESSING.
  ELSE.
    go_invloc_transfer->z_set_movtype_r7( i_dd_mtyp  = p_mtyp ).
  ENDIF.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

  IF p_file IS INITIAL.
    MESSAGE e008.                                " Please Select Valid File
    LEAVE LIST-PROCESSING.
  ENDIF.
  " Set file path to private member
  go_invloc_transfer->z_set_file_path( p_file ).

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
  " Removing Document date and Posting date from Selection Ccreen and defauliting the System date
  " Get data into internal table
*  go_invloc_transfer->z_get_data_into_table(
*                           i_doc_date  = p_doc
*                           i_pst_date  = p_post ).

  " Document date and postign date as system date: R7 requirement
  go_invloc_transfer->z_get_data_into_table(
                         i_doc_date  = sy-datum       " System date
                         i_pst_date  = sy-datum ).    " System date

* In R7 we've created new method 'z_set_movtype_r7', commenting out this 'z_set_mov_type'
*  "Set movement type
*  go_invloc_transfer->z_set_mov_type( i_rb_998  = p_rad1     " Radiobuttons are replated by Dropdown
*                                      i_rb_997  = p_rad2
*                                      i_rb_971  = p_rad3
*                                      i_rb_972  = p_rad4
*                                      i_rb_971w = p_rad5
*                                      i_rb_972w = p_rad6
*                                      i_rb_711  = p_rad7   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*                                      i_rb_712  = p_rad8   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*                                      i_rb_711w = p_rad9   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*                                      i_rb_712w = p_rad10  " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*                                     ).

* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

  "Validate the excel file
  go_invloc_transfer->z_validate_file( ).
  "Pass the filtered data from validation check to populate bapi tables
  IF gv_error EQ zif_gbl_constants=>gc_blank.
*    go_invloc_transfer->z_populate_bapi_tables( ).           "-- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    go_invloc_transfer->z_populate_bapi_tables_r7( ).       "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
  ENDIF.

END-OF-SELECTION.
* Display ALV

  go_invloc_transfer->z_display_alv( ).
