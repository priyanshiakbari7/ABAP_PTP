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
*04/22/2024  SDE2             R7-ORTHO  FTS.EXT.286         EC2K904625    *
* Replace the Movement Type Radiobuttons with Dropdown,                   *
* Dynamic download template and additional checks before posting          *
* Include 551W and 552W movementy types                                   *
*02/07/2025  SASHOK       R7 - Fix for WebGUI file operation EC1K923128   *
***************************************************************************

**********************************************************************
* SELECTION SCREEN Declaration
**********************************************************************

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
" Removing all movement type radio buttons from Selection Screen
" and replacing with dropdown list as per R7 design change

*SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-016.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad1  RADIOBUTTON GROUP gr1 DEFAULT 'X'. " Goods issue consignment - lending
*    SELECTION-SCREEN COMMENT 3(55) TEXT-028.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad2  RADIOBUTTON GROUP gr1.  " Goods issue consignment - return delivery
*    SELECTION-SCREEN COMMENT 3(55) TEXT-029.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad3  RADIOBUTTON GROUP gr1.  "Reversal of Entry of ST Balance - Own Stock
*    SELECTION-SCREEN COMMENT 3(65) TEXT-030.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad4 RADIOBUTTON GROUP gr1.  "Reversal of Entry of ST Balance - Customer Stock
*    SELECTION-SCREEN COMMENT 3(75) TEXT-031.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad5  RADIOBUTTON GROUP gr1.  " Initial Entry of ST Balance - Own Stock
*    SELECTION-SCREEN COMMENT 3(65) TEXT-032.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad6 RADIOBUTTON GROUP gr1.  "Initial Entry of ST Balance - Customer Stock
*    SELECTION-SCREEN COMMENT 3(65) TEXT-033.
*  SELECTION-SCREEN END OF LINE.
*
** Begin of Changes R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad7 RADIOBUTTON GROUP gr1.  "PI Adjustment - Decrease
*    SELECTION-SCREEN COMMENT 3(65) TEXT-034.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad8 RADIOBUTTON GROUP gr1.  "PI Adjustment - Increase
*    SELECTION-SCREEN COMMENT 3(65) TEXT-035.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad9 RADIOBUTTON GROUP gr1.  "PI Adjustment - Customer Stock Decrease
*    SELECTION-SCREEN COMMENT 3(65) TEXT-038.
*  SELECTION-SCREEN END OF LINE.
*
*  SELECTION-SCREEN BEGIN OF LINE.
*    PARAMETERS: p_rad10 RADIOBUTTON GROUP gr1.  "PI Adjustment - Customer Stock Increase
*    SELECTION-SCREEN COMMENT 3(65) TEXT-039.
*  SELECTION-SCREEN END OF LINE.
** End of Changes R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
*SELECTION-SCREEN END OF BLOCK b1.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
" Removing Document date and Posting date from Selection Screen and defauliting the System date
*SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
*  PARAMETERS: p_doc TYPE mkpf-bldat DEFAULT sy-datum.
*  PARAMETERS: p_post TYPE mkpf-budat DEFAULT sy-datum.
*SELECTION-SCREEN END OF BLOCK b2.

" Adding Movement type as Dropdown element on Selection Screen
SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-042.
  PARAMETERS: p_mtyp(4) TYPE c AS LISTBOX VISIBLE LENGTH 70 MODIF ID mov.
SELECTION-SCREEN END OF BLOCK b2.

* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-003.
  PARAMETERS: p_file TYPE rlgrap-filename. " Input file path

SELECTION-SCREEN END OF BLOCK b3.
**** Begin of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
  SELECTION-SCREEN BEGIN OF LINE.
    SELECTION-SCREEN COMMENT 1(79) gv_comm.
  SELECTION-SCREEN END OF LINE.
**** End of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
SELECTION-SCREEN: FUNCTION KEY 1.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
**********************************************************************
* AT SELECTION SCREEN OUTPUT
**********************************************************************
* Filling the list box
AT SELECTION-SCREEN OUTPUT.

  IF gt_vrm_values IS INITIAL.

    CLEAR gwa_vrm_value.
    APPEND gwa_vrm_value TO gt_vrm_values.

    "Fetch All relevant Entries from TVARVC Table
    DATA(lt_tvarvc_286) = zcl_otc_tvarv_utility=>z_get_tvarvc_tab_like(
                                EXPORTING i_variable_name = gc_tvarvc_286 ).

    "Get Maximum Posting Lines
    IF line_exists( lt_tvarvc_286[ name = gc_tvarvc_mpl ] ).
      gv_tvarv_mpl = lt_tvarvc_286[ name = gc_tvarvc_mpl
                                          type = zif_global_constants_n3=>gc_type_p ]-low.
    ENDIF.
    "Get Maximum Posting Value
    IF line_exists( lt_tvarvc_286[ name = gc_tvarvc_mpv ] ).
      gv_mpv = lt_tvarvc_286[ name = gc_tvarvc_mpv
                              type = zif_global_constants_n3=>gc_type_p ]-low.
    ENDIF.

    " Get the Movement Types
    SORT lt_tvarvc_286 BY low.
    DELETE lt_tvarvc_286 WHERE name NE gc_tvarvc_mtyp.
    LOOP AT lt_tvarvc_286 INTO DATA(ls_tvarvc_mtyp).
      gwa_vrm_value-key = ls_tvarvc_mtyp-low.
      gwa_vrm_value-text = ls_tvarvc_mtyp-high.
      APPEND gwa_vrm_value TO gt_vrm_values.
      CLEAR: ls_tvarvc_mtyp, gwa_vrm_value.
    ENDLOOP.

    CALL FUNCTION 'VRM_SET_VALUES'
      EXPORTING
        id     = gc_vrm_id
        values = gt_vrm_values.
*      EXCEPTIONS
*        id_illegavrm_id = 1
*        OTHERS          = 2.
*    IF sy-subrc <> 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
*    ENDIF.
  ENDIF.
**** Begin of insert R7-ORTHO/SASHOK/EC1K923128 For WebGUI
****  Display comment only for WEBGUI execution
  IF cl_gui_frontend_services=>www_active = abap_false.
    CLEAR gv_comm.
  ELSE.
    gv_comm = TEXT-m01.
  ENDIF.
**** End of insert R7-ORTHO/SASHOK/EC1K923128 For WebGUI
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

**********************************************************************
* AT SELECTION SCREEN ON VALUE REQUEST
**********************************************************************
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.

  DATA : gv_rc          TYPE i,
         gt_filetable_t TYPE STANDARD TABLE OF file_table.

* Methods for F4 help
  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      file_filter             = cl_gui_frontend_services=>filetype_excel
    CHANGING
      file_table              = gt_filetable_t    " Table Holding Selected Files
      rc                      = gv_rc           " Return Code, Number of Files or -1 If Error Occurred
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.

* Read and pass the value if file is selected
  IF sy-subrc IS INITIAL AND gv_rc = 1 .  " If only one file uploaded
    IF line_exists( gt_filetable_t[ 1 ] ).
      p_file = gt_filetable_t[ 1 ]-filename.
    ENDIF.
    CLEAR gv_rc.
  ENDIF.
