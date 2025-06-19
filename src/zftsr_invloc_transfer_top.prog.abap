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
*04/22/2024  SDE2             R7 ORTHO FTS.EXT.286          EC2K904625    *
* Replace the Movement Type Radiobuttons with Dropdown,                   *
* Dynamic download template and additional checks before posting          *
* Include 551W and 552W movementy types                                   *
*02/07/2025  SASHOK       R7 - Fix for WebGUI file operation EC1K923128   *
***************************************************************************

TABLES : sscrfields.
TYPES :
  BEGIN OF gty_bapi_data, "type for final internal table
    header_txt TYPE bktxt,
    plant      TYPE werks_d,
    stge_loc   TYPE lgort_d,
    customer   TYPE ekunn,
    material   TYPE matnr18,
    batch      TYPE charg_d,
    pstng_date TYPE budat,
    doc_date   TYPE bldat,
    entry_qnt  TYPE erfmg,
    entry_uom  TYPE erfme,
    serialno   TYPE gernr,
    move_type  TYPE bwart,
    stck_type  TYPE mb_insmk,
    spec_stock TYPE sobkz,
    prod_date  TYPE hsdat,
    expirydate TYPE vfdat,
    grund      TYPE mb_grbew, "R6-ENDO/FTS.EXT.069/CR#1299/VMANEM/EC2K902748
    mat_slip   TYPE mtsnr,    "R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
  END OF gty_bapi_data,

  BEGIN OF gty_mch1_data,
    matnr  TYPE matnr,
    charg  TYPE charg_d,
  END OF   gty_mch1_data.



DATA: gv_flag_postdata TYPE c.  "This flag will be set when POST is clicked
DATA: gv_flag_setstatus TYPE c. "This flag will be set in 2nd o/p screen
DATA: gv_error TYPE c.
DATA: gt_final TYPE STANDARD TABLE OF gty_bapi_data,
      gt_mch1  TYPE STANDARD TABLE OF gty_mch1_data,
      gt_mch2  TYPE STANDARD TABLE OF gty_mch1_data,
      wa_mch1  TYPE gty_mch1_data.

************************************************************************
* CLASS DEFINITION
************************************************************************

"! Class definition upload program and create sales order
CLASS lcl_invloc_transfer DEFINITION.

  PUBLIC SECTION.

    INTERFACES zif_gbl_constants.

* PUBLIC TYPE DECLARATION

    TYPES :

      BEGIN OF lty_excel_col, " Type for excel rows
        "! Document Header Text
        cola(40),
        "! plant
        colb(40),
        "! sloc
        colc(40),
        "! customer
        cold(40),
        "! material
        cole(40),
        "! batch
        colf(40),
        "! quantity
        colg(40),
        "! Unit of Entry
        colh(40),
        "!serial number
        coli(40),
        "! Movement Type: Double Check
        colj(40),
        "! Stock Type
        colk(40),
        "! Special Stock Indicator
        coll(40),
        "! prod date
        colm(40),
        "! expiry date
        coln(40),
        "! Reason code  " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        colo(40),       " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        "! Material Slip " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        colp(40),        " R7-ORTHO/sde2/fts.ext.286/ ec2k904625
      END OF lty_excel_col,

      BEGIN OF lty_exl_head_tab,  " Type For Template Heading table
        "! Document Header Text
        col1(40),
        "! plant
        col2(40),
        "! sloc
        col3(40),
        "! customer
        col4(40),
        "! material
        col5(40),
        "! batch
        col6(40),
        "! quantity
        col7(40),
        "! Unit of Entry
        col8(40),
        "!serial number
        col9(40),
        "! Movement Type: Double Check
        col10(40),
        "! Stock Type
        col11(40),
        "! Special Stock Indicator
        col12(40),
        "! prod date
        col13(40),
        "! expry date
        col14(40),
        "! Reason Code " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        col15(40),     " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        "! Material Slip " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        col16(40),       " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
      END OF lty_exl_head_tab,

      BEGIN OF lty_alv_output, " Type for ALV display
        "!Status Traffic lights
        status     TYPE char4,
        "!Header Text
        header_txt TYPE bktxt,
        "!plant
        plant      TYPE werks_d,
        "! Storage location
        sloc       TYPE lgort_d,
        "!customer
        customer   TYPE ekunn,
        "! material
        material   TYPE matnr,
        "!batch
        batch      TYPE charg_d,
        "!quantity
        quantity   TYPE char20,
        "!unit
        unit       TYPE erfme,
        "!serial number
        serialno   TYPE gernr,
        "!Movement Type
        move_type  TYPE bwart,
        "!Stock Type
        stck_type  TYPE mb_insmk,
        "!Special stock indicator
        spec_stock TYPE sobkz,
        "!Production date
        prod_date  TYPE hsdat,
        "!Expiry date
        expirydate TYPE vfdat,
        "!Reason code             " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        grund      TYPE mb_grbew, " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
        "! Material Slip          " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        mat_slip   TYPE mtsnr,    " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        "! Message from BAPI/Program
        message    TYPE bapi_msg,
      END OF lty_alv_output.

*Internal table
    DATA: lt_itab   TYPE STANDARD TABLE OF lcl_invloc_transfer=>lty_exl_head_tab. " for Excel Upload/Download Data

*Constants
    CONSTANTS: lc_max_count TYPE char3 VALUE '498'.

* PUBLIC METHOD DECLARATION

    METHODS:

      "! Constructor method
      constructor IMPORTING  i_filepath TYPE rlgrap-filename,

      "! Create table for Template Heading used in excel
      z_create_template_header_table,

      z_create_dyn_template_hdr_tab,                      " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      "! Methods to set the file path private variable
      z_set_file_path IMPORTING i_file_path TYPE rlgrap-filename,

      "! Methods to to get data from file into internal table
      z_get_data_into_table IMPORTING i_doc_date TYPE bldat
                                      i_pst_date TYPE budat,
      "! Methods to set movement type
      z_set_mov_type IMPORTING i_rb_998  TYPE c
                               i_rb_997  TYPE c
                               i_rb_971  TYPE c
                               i_rb_972  TYPE c
                               i_rb_971w TYPE c
                               i_rb_972w TYPE c
                               i_rb_711  TYPE c   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
                               i_rb_712  TYPE c   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
                               i_rb_711w TYPE c   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
                               i_rb_712w TYPE c,  " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748

      z_set_movtype_r7 IMPORTING i_dd_mtyp TYPE char4, " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      "!Method to validate excel file
      z_validate_file,
      "!Method to convert uom to internal unit
      z_convert_uom IMPORTING i_entry_uom TYPE erfme
                    EXPORTING e_entry_uom TYPE erfme,

      z_convert_date IMPORTING i_date TYPE sy-datum
                     EXPORTING e_date TYPE sy-datum,
      "!Method to pass all bapi data
      z_populate_bapi_tables,

      z_populate_bapi_tables_r7,           " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      z_call_bapi_goodsmvt_create,

      "! Method populate data in output table ( ALV Output)
      z_populate_alv_data
        IMPORTING i_v_status     TYPE char4
                  i_v_header     TYPE bktxt
                  i_v_plant      TYPE werks_d
                  i_v_sloc       TYPE lgort_d
                  i_v_customer   TYPE ekunn
                  i_v_material   TYPE matnr18
                  i_v_batch      TYPE charg_d
                  i_v_quantity   TYPE i
                  i_v_unit       TYPE erfme
                  i_v_serialno   TYPE gernr
                  i_v_move_type  TYPE bwart
                  i_v_stck_type  TYPE mb_insmk
                  i_v_spec_stock TYPE sobkz
                  i_v_prod_date  TYPE hsdat
                  i_v_expirydate TYPE vfdat
                  i_v_grund      TYPE mb_grbew " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
                  i_v_mat_slip   TYPE mtsnr    OPTIONAL  " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
                  i_v_message    TYPE bapi_msg,


      "!Method to populate serial data for bapi
      z_populate_bapi_serial_data
        IMPORTING i_wa_data TYPE gty_bapi_data
                  i_matid   TYPE mblpo,
      "! Method to display ALV
      z_display_alv,

      "! Method for user command
      z_on_click.

  PRIVATE SECTION.

* PRIVATE CLASS DATA DECLARATION
    DATA :
      "! File path
      lv_file_name    TYPE rlgrap-filename,
      "! Movement Type
      lv_move_type    TYPE bwart,
      "!Special stock indicator
      lv_spec_stock   TYPE sobkz,
      lv_update_batch TYPE c,
      lt_mch2         TYPE TABLE OF mch1,
      lv_count        TYPE i,     "counts number of line items in a material doc
      "!matid
      lv_matid        TYPE mblpo VALUE zif_gbl_constants=>gc_0001,
      "!Header work area
      lwa_header      TYPE bapi2017_gm_head_01,
      "! Item data for BAPI
      lt_items        TYPE TABLE OF bapi2017_gm_item_create,
      "! Serial no for BAPI
      lt_serial       TYPE TABLE OF bapi2017_gm_serialnumber,
      "! Excel data
      lt_data1        TYPE STANDARD TABLE OF lty_excel_col,
      "! ALV table which will be displayed as result
      lt_alv_output   TYPE STANDARD TABLE OF lty_alv_output.
ENDCLASS.

CLASS lcl_handle_events DEFINITION FINAL INHERITING FROM lcl_invloc_transfer.

  PUBLIC SECTION.

    METHODS z_on_user_command FOR EVENT added_function OF cl_salv_events.
ENDCLASS.                    "lcl_handle_events DEFINITION

* Create object ref
DATA: go_events TYPE REF TO lcl_handle_events.
DATA: go_events_alv TYPE REF TO cl_salv_events_table. " Events

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
TYPE-POOLS: vrm.

CONSTANTS: gc_alphabet(26) TYPE c          VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
           gc_tvarvc_mtyp  TYPE rvari_vnam VALUE 'FTS.EXT.286-MOVTYP',
           gc_vrm_id       TYPE vrm_id     VALUE 'P_MTYP'.

CONSTANTS: gc_tvarvc_mpl TYPE rvari_vnam VALUE 'FTS.EXT.286-MAX_POST_LINES',
           gc_tvarvc_mpv TYPE rvari_vnam VALUE 'FTS.EXT.286_MAX_POST_VALUE',
           gc_tvarvc_286 TYPE rvari_vnam VALUE 'FTS.EXT.286%'.

DATA: gt_vrm_values TYPE vrm_values,
      gwa_vrm_value LIKE LINE OF gt_vrm_values.

DATA: gv_mpv       TYPE p DECIMALS 3,
      gv_tvarv_mpl TYPE i.

DATA: gv_dynl TYPE i.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
