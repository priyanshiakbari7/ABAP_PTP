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
*06/02/2024  MBRUNO2          CHG0155111 / 8000014801       ECDK951030    *
* Adj To keep batch expiry date in cases it is blank in load file         *
*06/12/2024  MBRUNO2          CHG0155111 / 8000014923       ECDK951796    *
* Adj To keep batch expiry and manuf date in cases is blank in load file  *
*07/01/2024  MBRUNO2          CHG0158448 / 8000015138       ECDK953005    *
* Change logic on Batches with 00000000                                   *
*04/22/2024  SDE2             R7-ORTHO  FTS.EXT.286         EC2K904625    *
*                            CHG0136424 /  8000012486                     *
* Replace the Movement Type Radiobuttons with Dropdown,                   *
* Dynamic download template and additional checks before posting          *
* Include 551W and 552W movementy types                                   *
*12/10/2024  PAKBARI       R7-ORTHO FTS.EXT.286 - CR943443  EC1K923128    *
*         Update Reason code only for Capitalized batched managed materils*
*            SASHOK        R7 - Fix WebGUI file operation   EC1K923128    *
***************************************************************************

************************************************************************
* CLASS IMPLEMENTATION
************************************************************************

CLASS lcl_invloc_transfer IMPLEMENTATION.
**********************************************************************
* Method Description : Get the excel data in SAP internal table      *
**********************************************************************
  METHOD z_get_data_into_table.
*** BEGIN OF CHANGE CHG0155111/ECDK951796/MBRUNO2
    TYPES:
      BEGIN OF lty_batch,
        matnr TYPE mch1-matnr,
        charg TYPE mch1-charg,
      END OF lty_batch.

    DATA: lt_batch TYPE TABLE OF lty_batch.
*** END OF CHANGE CHG0155111/ECDK951796/MBRUNO2
**---->Begin of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
    TYPES:
      BEGIN OF lty_objek,
        objek TYPE cuobn,
      END OF lty_objek,

      BEGIN OF lty_custnum,
        custnum TYPE kunwe,
      END OF lty_custnum.
    DATA: lt_objek_tmp TYPE TABLE OF lty_objek,
          lt_custnum   TYPE TABLE OF lty_custnum,
          lv_atinn     TYPE atinn.
**---->End of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
    DATA : lv_doc_date TYPE bldat,
           lv_pst_date TYPE budat.

    DATA: lv_custnum TYPE kunnr,

*          lv_filename1 TYPE string, "++INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
          lv_rscode  TYPE zzaugru.          "++R7-ORTHO/MLANGALI/FTS.EXT.286/EC2K904625

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    CONSTANTS: lc_tvarvc_mpl       TYPE rvari_vnam VALUE 'FTS.EXT.286-MAX_POST_LINES',
               lc_tvarvc_mpv       TYPE rvari_vnam VALUE 'FTS.EXT.286_MAX_POST_VALUE',
               lc_tvarvc_286       TYPE rvari_vnam VALUE 'FTS.EXT.286%',
               lc_atnam            TYPE atnam      VALUE 'ZCAP_CATEGORIES',    "++R7-ORTHO/PAKBARI/CR943443/EC1K923128
               lc_atwrt            TYPE atwrt      VALUE 'Y_BATCH',            "++R7-ORTHO/PAKBARI/CR943443/EC1K923128
               lc_file_name_length TYPE i VALUE 255.
**** Begin of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
    CONSTANTS: lc_xls    TYPE char3 VALUE 'xls',
               lc_1      TYPE i VALUE '1',
               lc_row_01 TYPE i VALUE '01',
               lc_col_20 TYPE i VALUE '20'.
    DATA: lv_extn TYPE char10.
**** End of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
    DATA: lv_post_lines TYPE i.
    DATA: lv_mpv_flg TYPE boolean.
    DATA: lv_buom_flg  TYPE boolean,
          lv_filename1 TYPE c LENGTH lc_file_name_length.

    DATA: lv_matnr    TYPE mara-matnr,
          lv_inmenge  TYPE ekpo-menge,
          lv_outmenge TYPE ekpo-menge.

    DATA: lv_entry_uom  TYPE mara-meins,
          lv_mara_meins TYPE mara-meins.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    CLEAR: lt_data1,lv_doc_date,lv_pst_date.

    lv_doc_date = i_doc_date.
    lv_pst_date = i_pst_date.

    DATA(lo_file_upload) = NEW zcl_excel_file_upload( ).

    IF cl_gui_frontend_services=>www_active = abap_false.    "++INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
      lo_file_upload->upload_excel_file(
            EXPORTING
              i_filename          = lv_file_name
            IMPORTING
               c_tab_data         = lt_data1
            EXCEPTIONS
              conversion_failed   = 1 ).
**** Begin of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
    ELSE.
      lv_filename1 = lv_file_name.
      CALL FUNCTION 'TRINT_FILE_GET_EXTENSION'
        EXPORTING
          filename  = lv_filename1
        IMPORTING
          extension = lv_extn.
      lv_extn = to_lower( lv_extn ).
      IF lv_extn NE lc_xls.

        CALL METHOD lo_file_upload->z_upload_xlsx_file_webgui(
          EXPORTING
            iv_file                      = CONV string( lv_filename1 )
            iv_sheet_no                  = lc_1                  " Sheet Number 1
            iv_from_row                  = lc_row_01                " Start from Row 1
            iv_columns                   = lc_col_20                " Number of columns 20
          IMPORTING
            et_data                      = lt_data1
          EXCEPTIONS
            please_enter_filename        = 1
            sheet_no_missing             = 2
            row_no_missing               = 3
            column_no_missing            = 4
            invalid_worksheet_no         = 5
            no_worksheet_exists_in_excel = 6
            other                        = 7
            OTHERS                       = 8 ).

      ELSE.
        CALL METHOD lo_file_upload->z_upload_file_webgui
          EXPORTING
            i_file = CONV string( lv_filename1 )
          CHANGING
            c_tab  = lt_data1.

      ENDIF.
    ENDIF.
**** End of INS R7-ORTHO/SASHOK/EC1K923128 For WebGUI
    IF sy-subrc IS INITIAL.

      ASSIGN lt_data1[ 1 ] TO FIELD-SYMBOL(<lfs_data1>).
      IF sy-subrc IS INITIAL.
        " Basic check for template
* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        " Due to dynamic uplaod layout, Header Text and Plant is Mandatory accross all Mov_Types
*        IF <lfs_data1>-colb = TEXT-005 AND <lfs_data1>-colc = TEXT-006
*           AND <lfs_data1>-cold = TEXT-015 AND <lfs_data1>-cole = TEXT-007 .
        IF <lfs_data1>-cola = TEXT-004 AND <lfs_data1>-colb = TEXT-005.
          " Check at lease one record is there in template
*          IF line_exists( lt_data1[ 3 ] ).
*            DELETE lt_data1[]  FROM 1 TO 2.  " Delete the header
          " In new tempalte, ther eis only one Header Line
          IF line_exists( lt_data1[ 2 ] ).
            DELETE lt_data1[] INDEX 1.  " Delete the single header
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
          ELSE.
            MESSAGE  e005." Please enter at lease one record in excel
          ENDIF.
        ELSE.
          MESSAGE  e006." Excel template is not correct please download and fill the data
        ENDIF.
      ENDIF.
    ELSE.
      MESSAGE e007. " Error while reading data from Excel
    ENDIF.
    DATA: lv_proddate TYPE datum,
          lv_expdate  TYPE datum.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    IF NOT gv_tvarv_mpl IS INITIAL.
      DESCRIBE TABLE lt_data1 LINES lv_post_lines.
      IF lv_post_lines GT gv_tvarv_mpl.      " Check if the posting lines cross the max limit
        MESSAGE s042 WITH gv_tvarv_mpl DISPLAY LIKE zif_gbl_constants=>gc_e.
        LEAVE LIST-PROCESSING.
      ENDIF.
    ENDIF.
**---->Begin of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
    " Get Internal characteristic value for ZCAP_CATEGORIES
    CALL FUNCTION 'CONVERSION_EXIT_ATINN_INPUT'
      EXPORTING
        input  = lc_atnam
      IMPORTING
        output = lv_atinn.
    IF lt_data1 IS NOT INITIAL.
      " Preparing internal table with OBJEK as MATNR and COLC as CUSTNUM
      LOOP AT lt_data1 ASSIGNING FIELD-SYMBOL(<lfs_data_temp>).
        " Convert string COLC to proper format
        lv_custnum = |{ <lfs_data_temp>-colc(10) ALPHA = IN }|.
        APPEND lv_custnum TO lt_custnum.
        APPEND <lfs_data_temp>-cold TO lt_objek_tmp.
      ENDLOOP.
      IF lt_custnum IS NOT INITIAL.
        SORT lt_custnum BY custnum.
        DELETE ADJACENT DUPLICATES FROM lt_custnum.
        " Get Reason for Movement from ZFTS_RSCODEINFO
        SELECT custnum, rscode
          INTO TABLE @DATA(lt_rscode)
          FROM zfts_rscodeinfo
          FOR ALL ENTRIES IN @lt_custnum
        WHERE custnum = @lt_custnum-custnum
        AND   mvntype = @p_mtyp+0(3)
        AND   splind = @p_mtyp+3(1).
        IF sy-subrc = 0.
          SORT lt_rscode BY custnum.
        ENDIF.
      ENDIF.
      IF lt_objek_tmp IS NOT INITIAL.
        SORT lt_objek_tmp BY objek.
        DELETE ADJACENT DUPLICATES FROM lt_objek_tmp.
        " Get materials that are capitalized batch managed
        SELECT objek
          INTO TABLE @DATA(lt_objek)
          FROM ausp
          FOR ALL ENTRIES IN @lt_objek_tmp
        WHERE objek = @lt_objek_tmp-objek
        AND   atinn = @lv_atinn
        AND   atwrt = @lc_atwrt.
        IF sy-subrc = 0.
          SORT lt_objek BY objek.
        ENDIF.
      ENDIF.
    ENDIF.
**---->End of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
    CASE p_mtyp.
      WHEN zif_global_constants_n3=>gc_movement_711   " '711'
         OR zif_global_constants_n3=>gc_movement_712. " '712'.
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING FIELD-SYMBOL(<lfs_data>).
          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                              plant      = <lfs_data>-colb
                                              stge_loc   = <lfs_data>-colc
                                              material   = <lfs_data>-cold
                                              batch      = <lfs_data>-cole
                                              entry_qnt  = <lfs_data>-colf
                                              entry_uom  = <lfs_data>-colg
                                              serialno   = <lfs_data>-colh
                                              move_type  = <lfs_data>-coli
                                              stck_type  = <lfs_data>-colj ) ).

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cold
                                            charg   = <lfs_data>-cole ) ).
        ENDLOOP.

      WHEN zif_global_constants_n3=>gc_movement_711w      " '711W'
         OR zif_global_constants_n3=>gc_movement_712w .   " or '712W'
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING <lfs_data>.


**---->Begin of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
          READ TABLE lt_objek
            ASSIGNING FIELD-SYMBOL(<lfs_objek>)
            WITH KEY objek = <lfs_data>-cold
          BINARY SEARCH.
          " Process only if the material is capitalized batch managed
          IF sy-subrc = 0.
            CLEAR lv_custnum.
            lv_custnum = |{ <lfs_data>-colc(10) ALPHA = IN }|.
            " Update the Reason code based on ZFTS_RSCODEINFO table entry
            READ TABLE lt_rscode
              ASSIGNING FIELD-SYMBOL(<lfs_rscode>)
              WITH KEY custnum = lv_custnum
            BINARY SEARCH.
            IF sy-subrc = 0.
              <lfs_data>-colm = <lfs_rscode>-rscode.
            ENDIF.
          ENDIF.
**---->End of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                              plant      = <lfs_data>-colb
                                              customer   = <lfs_data>-colc
                                              material   = <lfs_data>-cold
                                              mat_slip   = <lfs_data>-cole   " Material Slip
                                              batch      = <lfs_data>-colf
                                              entry_qnt  = <lfs_data>-colg
                                              entry_uom  = <lfs_data>-colh
                                              serialno   = <lfs_data>-coli
                                              move_type  = <lfs_data>-colj
                                              stck_type  = <lfs_data>-colk
                                              spec_stock = <lfs_data>-coll
                                              grund      = <lfs_data>-colm ) ).

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cold
                                            charg   = <lfs_data>-cole ) ).
*          ENDIF.
        ENDLOOP.
      WHEN zif_gbl_constants=>gc_movement_971    " '971'
        OR zif_gbl_constants=>gc_movement_972.   "  OR '972'.
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING <lfs_data>.
          CLEAR: lv_proddate, lv_expdate.
          CONCATENATE: <lfs_data>-colk+6(4) <lfs_data>-colk+3(2) <lfs_data>-colk+0(2) INTO lv_proddate.
          CONCATENATE: <lfs_data>-coll+6(4) <lfs_data>-coll+3(2) <lfs_data>-coll+0(2) INTO lv_expdate.

          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                              plant      = <lfs_data>-colb
                                              stge_loc   = <lfs_data>-colc
                                              material   = <lfs_data>-cold
                                              batch      = <lfs_data>-cole
                                              entry_qnt  = <lfs_data>-colf
                                              entry_uom  = <lfs_data>-colg
                                              serialno   = <lfs_data>-colh
                                              move_type  = <lfs_data>-coli
                                              stck_type  = <lfs_data>-colj
                                              prod_date  = lv_proddate
                                              expirydate = lv_expdate ) ).

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cold
                                            charg   = <lfs_data>-cole ) ).
        ENDLOOP.

      WHEN zif_global_constants_n3=>gc_movement_971w   " '971W'
         OR zif_global_constants_n3=>gc_movement_972w. " OR '972W'
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING <lfs_data>.

          CLEAR: lv_proddate, lv_expdate, lv_custnum, lv_rscode.
          CONCATENATE: <lfs_data>-coll+6(4) <lfs_data>-coll+3(2) <lfs_data>-coll+0(2) INTO lv_proddate.
          CONCATENATE: <lfs_data>-colm+6(4) <lfs_data>-colm+3(2) <lfs_data>-colm+0(2) INTO lv_expdate.

**---->Begin of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
          READ TABLE lt_objek
            ASSIGNING <lfs_objek>
            WITH KEY objek = <lfs_data>-cold
          BINARY SEARCH.
          " Process only if the material is capitalized batch managed
          IF sy-subrc = 0.
            CLEAR lv_custnum.
            lv_custnum = |{ <lfs_data>-colc(10) ALPHA = IN }|.
            " Update the Reason code based on ZFTS_RSCODEINFO table entry
            READ TABLE lt_rscode
              ASSIGNING <lfs_rscode>
              WITH KEY custnum = lv_custnum
            BINARY SEARCH.
            IF sy-subrc = 0.
              lv_rscode = <lfs_rscode>-rscode.
            ENDIF.
          ENDIF.
**---->End of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                              plant      = <lfs_data>-colb
                                              customer   = <lfs_data>-colc
                                              material   = <lfs_data>-cold
                                              batch      = <lfs_data>-cole
                                              entry_qnt  = <lfs_data>-colf
                                              entry_uom  = <lfs_data>-colg
                                              serialno   = <lfs_data>-colh
                                              move_type  = <lfs_data>-coli
                                              stck_type  = <lfs_data>-colj
                                              spec_stock = <lfs_data>-colk
                                              prod_date  = lv_proddate
                                              expirydate = lv_expdate
                                              grund      = lv_rscode ) ).  "++R7-ORTHO/MLANGALI/FTS.EXT.286/EC2K906023

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cold
                                            charg      = <lfs_data>-cole ) ).
*          ENDIF.
        ENDLOOP.

      WHEN zif_gbl_constants=>gc_movement_997   " '997'
        OR zif_gbl_constants=>gc_movement_998.  " '998'
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING <lfs_data>.
          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                      plant      = <lfs_data>-colb
                                      stge_loc   = <lfs_data>-colc
                                      customer   = <lfs_data>-cold
                                      material   = <lfs_data>-cole
                                      batch      = <lfs_data>-colf
                                      entry_qnt  = <lfs_data>-colg
                                      entry_uom  = <lfs_data>-colh
                                      serialno   = <lfs_data>-coli
                                      move_type  = <lfs_data>-colj ) ).

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cole
                                            charg   = <lfs_data>-colf ) ).
        ENDLOOP.

      WHEN zif_global_constants_n3=>gc_movement_551w    " '551W'
         OR zif_global_constants_n3=>gc_movement_552w.  " OR '552W'.
        "fill final internal table from table containing excel file data
        LOOP AT lt_data1 ASSIGNING <lfs_data>.


**---->Begin of change R7-ORTHO/PAKBARI/CR943443/EC1K923128
          READ TABLE lt_objek
            ASSIGNING <lfs_objek>
            WITH KEY objek = <lfs_data>-cold
          BINARY SEARCH.
          " Process only if the material is capitalized batch managed
          IF sy-subrc = 0.
            CLEAR lv_custnum.
            lv_custnum = |{ <lfs_data>-colc(10) ALPHA = IN }|.
            " Update the Reason code based on ZFTS_RSCODEINFO table entry
            READ TABLE lt_rscode
              ASSIGNING <lfs_rscode>
              WITH KEY custnum = lv_custnum
            BINARY SEARCH.
            IF sy-subrc = 0.
              <lfs_data>-colm = <lfs_rscode>-rscode.
            ENDIF.
          ENDIF.
**---->End of change R7-ORTHO/PAKBARI/CR943443/EC1K923128

          gt_final = VALUE #( BASE gt_final ( header_txt = <lfs_data>-cola
                                      plant      = <lfs_data>-colb
                                      customer   = <lfs_data>-colc
                                      material   = <lfs_data>-cold
                                      batch      = <lfs_data>-colf
                                      entry_qnt  = <lfs_data>-colg
                                      entry_uom  = <lfs_data>-colh
                                      serialno   = <lfs_data>-coli
                                      move_type  = <lfs_data>-colj
                                      stck_type  = <lfs_data>-colk
                                      spec_stock = <lfs_data>-coll
                                      mat_slip   = <lfs_data>-cole   " Material Slip
                                      grund      = <lfs_data>-colm ) ).

          " Prepare separate table for materials
          gt_mch1 = VALUE #( BASE gt_mch1 ( matnr   = <lfs_data>-cold
                                            charg      = <lfs_data>-cole ) ).
*          ENDIF.
        ENDLOOP.
    ENDCASE.

    " Commenting the existing R6 logic
*    "!fill final internal table from table containing excel file data
*    LOOP AT lt_data1 ASSIGNING FIELD-SYMBOL(<lfs_data>).
*      CLEAR: lv_proddate, lv_expdate.
*      CONCATENATE: <lfs_data>-colm+6(4) <lfs_data>-colm+3(2) <lfs_data>-colm+0(2) INTO lv_proddate.
*      CONCATENATE: <lfs_data>-coln+6(4) <lfs_data>-coln+3(2) <lfs_data>-coln+0(2) INTO lv_expdate.
**lv_proddate = <lfs_data>-colm(of
*
*      gt_final = VALUE #( BASE gt_final (
*                                          header_txt = <lfs_data>-cola
*                                          plant      = <lfs_data>-colb
*                                          stge_loc   = <lfs_data>-colc
*                                          customer   = <lfs_data>-cold
*                                          material   = <lfs_data>-cole
*                                          batch      = <lfs_data>-colf
*                                          entry_qnt  = <lfs_data>-colg
*                                          entry_uom  = <lfs_data>-colh
*                                          serialno   = <lfs_data>-coli
*                                          move_type  = <lfs_data>-colj
*                                          stck_type  = <lfs_data>-colk
*                                          spec_stock = <lfs_data>-coll
*                                          prod_date  = lv_proddate
*                                          expirydate = lv_expdate
*                                          grund      = <lfs_data>-colo " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
*                                          ) ).
*    ENDLOOP.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

*    loop at gt_final ASSIGNING FIELD-SYMBOL(<lfs_final1>).
*      z_convert_date( EXPORTING i_date = <lfs_final1>-prod_date
*                      IMPORTING e_date = <lfs_final1>-prod_date ).
*
*      z_convert_date( EXPORTING i_date = <lfs_final1>-expirydate
*                      IMPORTING e_date = <lfs_final1>-expirydate ).
*      endloop.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    IF NOT gt_mch1 IS INITIAL.
      SORT gt_mch1 BY matnr.
      DELETE ADJACENT DUPLICATES FROM gt_mch1 COMPARING matnr.
      " Get Base UoM for all unique material entries
      SELECT matnr , meins FROM mara
        INTO TABLE @DATA(lt_mara)
        FOR ALL ENTRIES IN @gt_mch1
        WHERE matnr = @gt_mch1-matnr.
    ENDIF.
    CLEAR: gt_mch1[].
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    LOOP AT gt_final ASSIGNING FIELD-SYMBOL(<lfs_final>).
      <lfs_final>-doc_date = lv_doc_date.
      <lfs_final>-pstng_date = lv_pst_date.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
      " Get Material Base Uom
      IF line_exists( lt_mara[ matnr = <lfs_final>-material ] ).
        lv_mara_meins = lt_mara[ matnr = <lfs_final>-material ]-meins.

        IF <lfs_final>-entry_uom NE lv_mara_meins.  " Compare Base Uom & Upload UoM

          lv_matnr   = <lfs_final>-material.
          lv_inmenge = <lfs_final>-entry_qnt.
          lv_entry_uom = <lfs_final>-entry_uom.

          CALL FUNCTION 'CONVERSION_EXIT_CUNIT_INPUT'
            EXPORTING
              input          = lv_entry_uom
              language       = sy-langu
            IMPORTING
              output         = lv_entry_uom
            EXCEPTIONS
              unit_not_found = 1
              OTHERS         = 2.
          IF sy-subrc <> 0.
* Implement suitable error handling here
          ENDIF.

          CALL FUNCTION 'MD_CONVERT_MATERIAL_UNIT'   " Convert into Base UoM
            EXPORTING
              i_matnr              = lv_matnr
              i_in_me              = lv_entry_uom
              i_out_me             = lv_mara_meins
              i_menge              = lv_inmenge
            IMPORTING
              e_menge              = lv_outmenge
            EXCEPTIONS
              error_in_application = 1
              error                = 2
              OTHERS               = 3.
          IF sy-subrc NE 0.
            lv_buom_flg = abap_true.
            EXIT.
          ELSE.
            <lfs_final>-entry_qnt = lv_outmenge.
            <lfs_final>-entry_uom = lv_mara_meins.   "Set Upload UoM as Base UoM
          ENDIF.
        ENDIF.
      ENDIF.

      " Check if the posting value cross the max limit
      IF NOT gv_mpv IS INITIAL.
        IF <lfs_final>-entry_qnt GT gv_mpv.
          lv_mpv_flg = abap_true.    " Set flag
          EXIT.
        ENDIF.
      ENDIF.

      CLEAR: lv_matnr, lv_inmenge, lv_outmenge, lv_entry_uom , lv_entry_uom .
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    ENDLOOP.

*** BEGIN OF CHANGE CHG0155111/ECDK951796/MBRUNO2
    IF gt_final IS NOT INITIAL.
      FREE lt_batch.
      LOOP AT gt_final ASSIGNING <lfs_final>.
        APPEND VALUE lty_batch( matnr = <lfs_final>-material
                                charg = <lfs_final>-batch ) TO lt_batch.
      ENDLOOP.
      SORT lt_batch BY matnr charg.
      DELETE ADJACENT DUPLICATES FROM lt_batch COMPARING matnr charg.

*** BEGIN OF CHANGE CHG0158448/ECDK953005/MBRUNO2
      DELETE lt_batch WHERE charg = zif_gbl_constants=>gc_0.
      IF lt_batch IS NOT INITIAL.
*** END OF CHANGE CHG0158448/ECDK953005/MBRUNO2
        SELECT matnr, charg, hsdat, vfdat
          FROM mch1
          INTO TABLE @DATA(lt_mch1_old)
          FOR ALL ENTRIES IN @lt_batch
          WHERE matnr EQ @lt_batch-matnr
            AND charg EQ @lt_batch-charg.
        SORT lt_mch1_old BY matnr charg.

        " If batch exists keep DOM and Expiry Date as it is in MCH1 table
        LOOP AT gt_final ASSIGNING <lfs_final>.
          READ TABLE lt_mch1_old
            INTO DATA(ls_mch1_old)
            WITH KEY matnr = <lfs_final>-material
                     charg = <lfs_final>-batch
          BINARY SEARCH.

          IF sy-subrc IS INITIAL.
            <lfs_final>-expirydate = ls_mch1_old-vfdat.
            <lfs_final>-prod_date  = ls_mch1_old-hsdat.
          ENDIF.
        ENDLOOP.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        " if the any posting value cross the max limit
        IF lv_mpv_flg = abap_true.
          MESSAGE s043 WITH gv_mpv DISPLAY LIKE zif_gbl_constants=>gc_e.
          LEAVE LIST-PROCESSING.
        ENDIF.

        IF lv_buom_flg = abap_true.
          MESSAGE s044 DISPLAY LIKE zif_gbl_constants=>gc_e.
          LEAVE LIST-PROCESSING.
        ENDIF.

* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

*** BEGIN OF CHANGE CHG0158448/ECDK953005/MBRUNO2
      ENDIF.
*** END OF CHANGE CHG0158448/ECDK953005/MBRUNO2
    ENDIF.
*** END OF CHANGE CHG0155111/ECDK951796/MBRUNO2

  ENDMETHOD.
**********************************************************************
* Method Description : Method to validate excel file data and        *
*                        populate alv table                          *
**********************************************************************
  METHOD z_populate_bapi_tables.
    DATA: lv_qnt           TYPE erfmg,
          lv_new_gdsmvt    TYPE c,
          lwa_data         TYPE gty_bapi_data,
          ls_goodsmvt_item TYPE bapi2017_gm_item_create,
          lv_new_item      TYPE c.

    CLEAR: lv_count, lv_new_item.

    "sort table to identify each set of material document that has to be created
    SORT gt_final BY plant stge_loc customer material batch.

    LOOP AT gt_final INTO DATA(wa_final).

      wa_final-customer = |{ wa_final-customer ALPHA = IN }|.

      z_convert_uom( EXPORTING i_entry_uom = wa_final-entry_uom
                     IMPORTING e_entry_uom = wa_final-entry_uom ).

      z_convert_date( EXPORTING i_date = wa_final-prod_date
                      IMPORTING e_date = wa_final-prod_date ).

      z_convert_date( EXPORTING i_date = wa_final-expirydate
                      IMPORTING e_date = wa_final-expirydate ).

      lv_qnt = lv_qnt + wa_final-entry_qnt.
      lv_count = lv_count + 1.
      lwa_data = CORRESPONDING #( wa_final ).
      AT NEW customer.
        lv_new_gdsmvt = zif_gbl_constants~gc_x.
      ENDAT.

      lwa_header = CORRESPONDING #( lwa_data ).

      lwa_header-ref_doc_no_long = lwa_data-mat_slip. " Material Slip " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      ls_goodsmvt_item = CORRESPONDING #( lwa_data EXCEPT entry_qnt ).

      ls_goodsmvt_item-entry_qnt = lv_qnt.
* Reason Code
      ls_goodsmvt_item-move_reas = wa_final-grund.    " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
      "Populate Serial data
      z_populate_bapi_serial_data( i_wa_data = lwa_data
                                   i_matid = lv_matid ).
      AT END OF batch.
        lv_matid   = lv_matid + 1.
        lv_new_item = zif_gbl_constants~gc_x.

      ENDAT.

      AT END OF customer.
        CLEAR: lv_new_gdsmvt, lv_new_item.
        MOVE '0001' TO lv_matid.
      ENDAT.

      IF lv_new_item = zif_gbl_constants~gc_x.
        APPEND ls_goodsmvt_item TO lt_items.
        CLEAR: lv_qnt.
      ENDIF.

      IF lv_count > lc_max_count OR lv_new_gdsmvt IS INITIAL.
        "Append Item data
        APPEND ls_goodsmvt_item TO lt_items.
        z_call_bapi_goodsmvt_create( ).
        CLEAR: lv_count, lv_qnt.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

* Begin of Change: R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
* Method z_populate_bapi_tables will be replaced by z_populate_bapi_tables_r7
**********************************************************************
* Method Description : Method to validate excel file data and        *
*                      populate alv table without Quantity Clubbing  *
**********************************************************************
  METHOD z_populate_bapi_tables_r7.
    DATA: lwa_data         TYPE gty_bapi_data,
          ls_goodsmvt_item TYPE bapi2017_gm_item_create,
          lv_new_item      TYPE c.

    "sort table to identify each set of material document that has to be created
    SORT gt_final BY plant stge_loc customer material batch.

    LOOP AT gt_final INTO DATA(wa_final).

      wa_final-customer = |{ wa_final-customer ALPHA = IN }|.

      z_convert_uom( EXPORTING i_entry_uom = wa_final-entry_uom
                     IMPORTING e_entry_uom = wa_final-entry_uom ).

      z_convert_date( EXPORTING i_date = wa_final-prod_date
                      IMPORTING e_date = wa_final-prod_date ).

      z_convert_date( EXPORTING i_date = wa_final-expirydate
                      IMPORTING e_date = wa_final-expirydate ).

      lwa_data = CORRESPONDING #( wa_final ).

      lwa_header = CORRESPONDING #( lwa_data ).

      lwa_header-ref_doc_no_long = lwa_data-mat_slip. " Material Slip

      ls_goodsmvt_item = CORRESPONDING #( lwa_data ).

      ls_goodsmvt_item-move_reas = wa_final-grund.    "  Reason Code

      " Populate Serial data
      z_populate_bapi_serial_data( i_wa_data = lwa_data
                                   i_matid = lv_matid ).
      "Append Item data
      APPEND ls_goodsmvt_item TO lt_items.
      z_call_bapi_goodsmvt_create( ).

    ENDLOOP.

  ENDMETHOD.
* End of Change: R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

**********************************************************************
* Method Description : Method to convert date to                     *
*                        internal unit                               *
**********************************************************************
  METHOD z_convert_date.
    CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
      EXPORTING
        date_external            = i_date  " external date formatting
      IMPORTING
        date_internal            = e_date  " internal date formatting
      EXCEPTIONS
        date_external_is_invalid = 1
        OTHERS                   = 2.
    IF sy-subrc IS NOT INITIAL.
      sy-subrc = 0.
    ENDIF.
  ENDMETHOD.

**********************************************************************
* Method Description : Method to convert commercial unit to          *
*                        internal unit                               *
**********************************************************************
  METHOD z_convert_uom.

    CALL FUNCTION 'CONVERSION_EXIT_CUNIT_INPUT'
      EXPORTING
        input          = i_entry_uom
        language       = sy-langu
      IMPORTING
        output         = e_entry_uom
      EXCEPTIONS
        unit_not_found = 1
        OTHERS         = 2.
    IF sy-subrc IS NOT INITIAL.
      sy-subrc = 0.
    ENDIF.
  ENDMETHOD.
**********************************************************************
* Method Description : Method to validate excel file data and        *
*                        populate alv table                          *
**********************************************************************
  METHOD z_validate_file.
    DATA: lv_qnt     TYPE i.
    DATA: lv_matnr40 TYPE matnr40.

*** BEGIN OF CHANGE CHG0158448/ECDK953005/MBRUNO2
    DATA: lr_matnr TYPE RANGE OF matnr.

    IF gt_final IS NOT INITIAL.
      LOOP AT gt_final INTO DATA(ls_final).
        APPEND VALUE #( sign   = zif_gbl_constants=>gc_i
                        option = zif_gbl_constants=>gc_eq
                        low    = ls_final-material ) TO lr_matnr.
      ENDLOOP.

      SELECT matnr, mhdhb
        INTO TABLE @DATA(lt_mara)
        FROM mara
        FOR ALL ENTRIES IN @lr_matnr
        WHERE matnr EQ @lr_matnr-low.
      SORT lt_mara BY matnr.
    ENDIF.
*** END OF CHANGE CHG0158448/ECDK953005/MBRUNO2
*** BEGIN OF CHANGE CHG0155111/ECDK951796/MBRUNO2
    " Following code already existed, it was moved to be beginning of form
    LOOP AT gt_final ASSIGNING FIELD-SYMBOL(<lfs_final_s>).
      gt_mch1 = VALUE #( BASE gt_mch1  ( matnr = <lfs_final_s>-material
                                         charg = <lfs_final_s>-batch ) ).
    ENDLOOP.

    IF gt_mch1 IS NOT INITIAL.
      SELECT matnr charg FROM mch1 INTO TABLE gt_mch2
                         FOR ALL ENTRIES IN gt_mch1
                         WHERE matnr EQ gt_mch1-matnr
                         AND   charg EQ gt_mch1-charg.
      IF sy-subrc IS INITIAL.
        SORT gt_mch2 BY matnr charg.
      ENDIF.
    ENDIF.
*** END OF CHANGE CHG0155111/ECDK951796/MBRUNO2

    LOOP AT gt_final INTO DATA(wa_data).
      lv_matnr40 = wa_data-material.
      lv_qnt = wa_data-entry_qnt.

      z_convert_uom( EXPORTING i_entry_uom = wa_data-entry_uom
                    IMPORTING e_entry_uom = wa_data-entry_uom ).   " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      "If 5th or 6th radiobutton selected set special stock indicator as w.
      IF lv_spec_stock EQ zif_gbl_constants=>gc_x.
        IF wa_data-spec_stock EQ ' '.
          CALL METHOD z_populate_alv_data
            EXPORTING
              i_v_status     = zif_gbl_constants=>gc_red
              i_v_header     = wa_data-header_txt
              i_v_plant      = wa_data-plant
              i_v_sloc       = wa_data-stge_loc
              i_v_customer   = wa_data-customer
              i_v_material   = wa_data-material
              i_v_batch      = wa_data-batch
              i_v_quantity   = lv_qnt
              i_v_unit       = wa_data-entry_uom
              i_v_serialno   = wa_data-serialno
              i_v_move_type  = wa_data-move_type
              i_v_stck_type  = wa_data-stck_type
              i_v_spec_stock = wa_data-spec_stock
              i_v_prod_date  = wa_data-prod_date
              i_v_expirydate = wa_data-expirydate
              i_v_grund      = wa_data-grund " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
              i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
              i_v_message    = TEXT-022.     "special stock indicator should be w
          gv_error = abap_true.
        ENDIF.
      ENDIF.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
      " For the Condition: Quantity in upload file exceeds threshold value
      IF NOT gv_mpv IS INITIAL
         AND lv_qnt GT gv_mpv.

        CALL METHOD z_populate_alv_data
          EXPORTING
            i_v_status     = zif_gbl_constants=>gc_red
            i_v_header     = wa_data-header_txt
            i_v_plant      = wa_data-plant
            i_v_sloc       = wa_data-stge_loc
            i_v_customer   = wa_data-customer
            i_v_material   = wa_data-material
            i_v_batch      = wa_data-batch
            i_v_quantity   = lv_qnt
            i_v_unit       = wa_data-entry_uom
            i_v_serialno   = wa_data-serialno
            i_v_move_type  = wa_data-move_type
            i_v_stck_type  = wa_data-stck_type
            i_v_spec_stock = wa_data-spec_stock
            i_v_prod_date  = wa_data-prod_date
            i_v_expirydate = wa_data-expirydate
            i_v_grund      = wa_data-grund
            i_v_mat_slip   = wa_data-mat_slip
            i_v_message    = TEXT-043.  " 'Quantity in upload file exceeds threshold value'.
        gv_error = abap_true.
      ENDIF.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

      "Validate the movement type in excel file based on the selection made for upload file
      IF wa_data-move_type = lv_move_type.
      ELSE.

        CALL METHOD z_populate_alv_data
          EXPORTING
            i_v_status     = zif_gbl_constants=>gc_red
            i_v_header     = wa_data-header_txt
            i_v_plant      = wa_data-plant
            i_v_sloc       = wa_data-stge_loc
            i_v_customer   = wa_data-customer
            i_v_material   = wa_data-material
            i_v_batch      = wa_data-batch
            i_v_quantity   = lv_qnt
            i_v_unit       = wa_data-entry_uom
            i_v_serialno   = wa_data-serialno
            i_v_move_type  = wa_data-move_type
            i_v_stck_type  = wa_data-stck_type
            i_v_spec_stock = wa_data-spec_stock
            i_v_prod_date  = wa_data-prod_date
            i_v_expirydate = wa_data-expirydate
            i_v_grund      = wa_data-grund " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
            i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
            i_v_message    = TEXT-024.     "movement type in excel file should be same as radio selection
        gv_error = abap_true.
      ENDIF.
      "! If the user has selected 3rd or 4th radiobutton
      "! validate the upload file so that the customer number is not entered
      IF ( lv_move_type EQ zif_gbl_constants=>gc_movement_971 OR
        lv_move_type EQ zif_gbl_constants=>gc_movement_972  OR
        lv_move_type EQ zif_global_constants_n3=>gc_movement_711 OR   " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
        lv_move_type EQ zif_global_constants_n3=>gc_movement_712      " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
        ) AND
        lv_spec_stock EQ zif_gbl_constants=>gc_blank.
        IF wa_data-customer NE zif_gbl_constants=>gc_blank.
          CALL METHOD z_populate_alv_data
            EXPORTING
              i_v_status     = zif_gbl_constants=>gc_red
              i_v_header     = wa_data-header_txt
              i_v_plant      = wa_data-plant
              i_v_sloc       = wa_data-stge_loc
              i_v_customer   = wa_data-customer
              i_v_material   = wa_data-material
              i_v_batch      = wa_data-batch
              i_v_quantity   = lv_qnt
              i_v_unit       = wa_data-entry_uom
              i_v_serialno   = wa_data-serialno
              i_v_move_type  = wa_data-move_type
              i_v_stck_type  = wa_data-stck_type
              i_v_spec_stock = wa_data-spec_stock
              i_v_prod_date  = wa_data-prod_date
              i_v_expirydate = wa_data-expirydate
              i_v_grund      = wa_data-grund " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
              i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
              i_v_message    = TEXT-023.     "As upload is for plant stock,do not enter customer
          gv_error = abap_true.
        ENDIF.
      ELSE.
* Begin of Changes for R6-ENDO/BUG 669327/VMANEM/EC1K916745
* Storage location value is not needed, when posting is for customer consignment stock.
        IF ( lv_move_type EQ zif_gbl_constants=>gc_movement_971 OR
           lv_move_type EQ zif_gbl_constants=>gc_movement_972  OR
           lv_move_type EQ zif_global_constants_n3=>gc_movement_711 OR
           lv_move_type EQ zif_global_constants_n3=>gc_movement_712
           ) AND
           lv_spec_stock NE zif_gbl_constants=>gc_blank AND
           wa_data-stge_loc IS NOT INITIAL AND
           wa_data-spec_stock EQ zif_gbl_constants=>gc_w.
          CALL METHOD z_populate_alv_data
            EXPORTING
              i_v_status     = zif_gbl_constants=>gc_red
              i_v_header     = wa_data-header_txt
              i_v_plant      = wa_data-plant
              i_v_sloc       = wa_data-stge_loc
              i_v_customer   = wa_data-customer
              i_v_material   = wa_data-material
              i_v_batch      = wa_data-batch
              i_v_quantity   = lv_qnt
              i_v_unit       = wa_data-entry_uom
              i_v_serialno   = wa_data-serialno
              i_v_move_type  = wa_data-move_type
              i_v_stck_type  = wa_data-stck_type
              i_v_spec_stock = wa_data-spec_stock
              i_v_prod_date  = wa_data-prod_date
              i_v_expirydate = wa_data-expirydate
              i_v_grund      = wa_data-grund " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
              i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
              i_v_message    = TEXT-041.     "As upload is for plant stock,do not enter customer
          gv_error = abap_true.
        ENDIF.
* End of Changes for R6-ENDO/BUG 669327/VMANEM/EC1K916745
      ENDIF.

*** BEGIN OF CHANGE CHG0155111/ECDK951796/MBRUNO2
      " Determine if DOM and SLED must be required
      DATA(lv_check_batch_date) = CONV abap_bool( abap_false ).

      IF wa_data-batch IS NOT INITIAL.
        READ TABLE gt_mch2
          TRANSPORTING NO FIELDS
          WITH KEY matnr = lv_matnr40
                   charg = wa_data-batch
        BINARY SEARCH.
        lv_check_batch_date = COND abap_bool( WHEN sy-subrc IS INITIAL THEN abap_false ELSE abap_true ).
      ENDIF.

*** BEGIN OF CHANGE CHG0158448/ECDK953005/MBRUNO2
      DATA(lv_validate_dates) = CONV abap_bool( abap_true ).

      IF wa_data-batch IS INITIAL OR
         wa_data-batch EQ zif_gbl_constants=>gc_0.

        lv_validate_dates = abap_false.
      ELSE.
        READ TABLE lt_mara
          INTO DATA(ls_mara)
          WITH KEY matnr = wa_data-material
        BINARY SEARCH.

        IF sy-subrc IS INITIAL AND
           ls_mara-mhdhb = zif_gbl_constants=>gc_0.

          lv_validate_dates = abap_false.
        ENDIF.
      ENDIF.

      IF lv_validate_dates EQ abap_true.
*** END OF CHANGE CHG0158448/ECDK953005/MBRUNO2
        " Check if manufacture date is greater than expiry date
        IF wa_data-prod_date GT wa_data-expirydate.
          CALL METHOD z_populate_alv_data
            EXPORTING
              i_v_status     = zif_gbl_constants=>gc_red
              i_v_header     = wa_data-header_txt
              i_v_plant      = wa_data-plant
              i_v_sloc       = wa_data-stge_loc
              i_v_customer   = wa_data-customer
              i_v_material   = wa_data-material
              i_v_batch      = wa_data-batch
              i_v_quantity   = lv_qnt
              i_v_unit       = wa_data-entry_uom
              i_v_serialno   = wa_data-serialno
              i_v_move_type  = wa_data-move_type
              i_v_stck_type  = wa_data-stck_type
              i_v_spec_stock = wa_data-spec_stock
              i_v_prod_date  = wa_data-prod_date
              i_v_expirydate = wa_data-expirydate
              i_v_grund      = wa_data-grund
              i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
              i_v_message    = TEXT-044.     "DOM cannot be greather than SLED
          gv_error = abap_true.
        ELSEIF lv_check_batch_date EQ abap_true AND
               ( wa_data-prod_date   EQ zif_gbl_constants=>gc_blank OR
                 wa_data-expirydate  EQ zif_gbl_constants=>gc_blank ).

          " If is a new batch, date fields must be provided
          CALL METHOD z_populate_alv_data
            EXPORTING
              i_v_status     = zif_gbl_constants=>gc_red
              i_v_header     = wa_data-header_txt
              i_v_plant      = wa_data-plant
              i_v_sloc       = wa_data-stge_loc
              i_v_customer   = wa_data-customer
              i_v_material   = wa_data-material
              i_v_batch      = wa_data-batch
              i_v_quantity   = lv_qnt
              i_v_unit       = wa_data-entry_uom
              i_v_serialno   = wa_data-serialno
              i_v_move_type  = wa_data-move_type
              i_v_stck_type  = wa_data-stck_type
              i_v_spec_stock = wa_data-spec_stock
              i_v_prod_date  = wa_data-prod_date
              i_v_expirydate = wa_data-expirydate
              i_v_grund      = wa_data-grund
              i_v_mat_slip   = wa_data-mat_slip  " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
              i_v_message    = TEXT-045.     "For new Batch DOM and SLED are required

          gv_error = abap_true.
        ENDIF.
*** BEGIN OF CHANGE CHG0158448/ECDK953005/MBRUNO2
      ENDIF.
*** END OF CHANGE CHG0158448/ECDK953005/MBRUNO2
*** END OF CHANGE CHG0155111/ECDK951796/MBRUNO2

      "check material batch exits, if yes ignore DOM & SLED

      IF wa_data-prod_date NE zif_gbl_constants=>gc_blank OR
        wa_data-expirydate NE zif_gbl_constants=>gc_blank.
        SELECT COUNT( * ) FROM mch1 UP TO 1 ROWS
          WHERE matnr EQ @lv_matnr40
            AND charg EQ @wa_data-batch.
        IF sy-subrc IS INITIAL.
          CLEAR: wa_data-prod_date, wa_data-expirydate.
          lv_update_batch = zif_gbl_constants~gc_x.
        ENDIF.
      ENDIF.
    ENDLOOP.

*** BEGIN OF COMMENT CHG0155111/ECDK951796/MBRUNO2
    " MOVED TO THE BEGINING OF FORM
*    LOOP AT gt_final ASSIGNING FIELD-SYMBOL(<lfs_final_s>).
*      gt_mch1 = VALUE #( BASE gt_mch1  ( matnr = <lfs_final_s>-material
*                                         charg = <lfs_final_s>-batch ) ).
*    ENDLOOP.
*
*    IF gt_mch1 IS NOT INITIAL.
*      SELECT matnr charg FROM mch1 INTO TABLE gt_mch2
*                         FOR ALL ENTRIES IN gt_mch1
*                         WHERE matnr EQ gt_mch1-matnr
*                         AND   charg EQ gt_mch1-charg.
*      IF sy-subrc IS INITIAL.
*        SORT gt_mch2 BY matnr charg.
*      ENDIF.
*    ENDIF.
*** END OF COMMENT CHG0155111/ECDK951796/MBRUNO2
  ENDMETHOD.
**********************************************************************
* Method Description : Method to set movement type                   *
**********************************************************************
  METHOD z_set_mov_type.
    IF i_rb_998     = zif_gbl_constants=>gc_x.
      lv_move_type  = zif_gbl_constants=>gc_movement_998.
    ELSEIF i_rb_997 = zif_gbl_constants=>gc_x.
      lv_move_type  = zif_gbl_constants=>gc_movement_997.
    ELSEIF i_rb_971 = zif_gbl_constants=>gc_x.
      lv_move_type  = zif_gbl_constants=>gc_movement_971.
    ELSEIF i_rb_972 = zif_gbl_constants=>gc_x.
      lv_move_type  = zif_gbl_constants=>gc_movement_972.
    ELSEIF i_rb_971w = zif_gbl_constants=>gc_x.
      lv_move_type   = zif_gbl_constants=>gc_movement_971.
      lv_spec_stock  = zif_gbl_constants=>gc_x.
    ELSEIF i_rb_972w = zif_gbl_constants=>gc_x.
      lv_move_type   = zif_gbl_constants=>gc_movement_972.
      lv_spec_stock  = zif_gbl_constants=>gc_x.
* Begin of Changes R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
    ELSEIF i_rb_711  = zif_gbl_constants=>gc_x..
      lv_move_type   = zif_global_constants_n3=>gc_movement_711.
    ELSEIF i_rb_712  = zif_gbl_constants=>gc_x.
      lv_move_type   = zif_global_constants_n3=>gc_movement_712.
    ELSEIF i_rb_711w = zif_gbl_constants=>gc_x.
      lv_move_type   = zif_global_constants_n3=>gc_movement_711.
      lv_spec_stock  = zif_gbl_constants=>gc_x.
    ELSEIF i_rb_712w = zif_gbl_constants=>gc_x.
      lv_move_type   = zif_global_constants_n3=>gc_movement_712.
      lv_spec_stock  = zif_gbl_constants=>gc_x.
* End of Changes R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
    ENDIF.
  ENDMETHOD.
**********************************************************************
* Method Description : Method for calling BAPI                       *
**********************************************************************
  METHOD z_call_bapi_goodsmvt_create.

    " LOCAL DATA DECLARATION
    DATA: lt_return           TYPE TABLE OF bapiret2,
          lt_mch1             TYPE TABLE OF mch1,
          lv_disp_field       TYPE c,
          ls_goodsmvt_code    TYPE bapi2017_gm_code,
          ls_return           TYPE bapiret2,
          ls_materialdocument TYPE bapi2017_gm_head_ret-mat_doc,
          ls_materialdocyear  TYPE bapi2017_gm_head_ret-doc_year,
          lv_testrun          TYPE c VALUE zif_gbl_constants=>gc_x,
          lv_message          TYPE bapi_msg,
          lv_message1         TYPE bapi_msg,
          lv_status           TYPE char4,
          lv_quantity         TYPE i.

    DATA: lv_mtsnr TYPE mtsnr.   "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    "!This flag will be set when POST is clicked
    IF gv_flag_postdata = zif_gbl_constants=>gc_x.
      lv_testrun = zif_gbl_constants=>gc_blank.
    ENDIF.

    "Assign code to transaction for BAPI goods movement
    ls_goodsmvt_code = zif_gbl_constants=>gc_06.

    CALL FUNCTION 'BAPI_GOODSMVT_CREATE'
      EXPORTING
        goodsmvt_header       = lwa_header
        goodsmvt_code         = ls_goodsmvt_code
        testrun               = lv_testrun
      IMPORTING
        materialdocument      = ls_materialdocument
        matdocumentyear       = ls_materialdocyear
      TABLES
        goodsmvt_item         = lt_items
        goodsmvt_serialnumber = lt_serial
        return                = lt_return.
    IF sy-subrc IS INITIAL.
      "!This flag will be set when POST is clicked
      IF gv_flag_postdata = zif_gbl_constants=>gc_x.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait   = zif_gbl_constants=>gc_x
          IMPORTING
            return = ls_return.
        IF sy-subrc IS INITIAL.
          sy-subrc = 0.
        ENDIF.
      ENDIF.
    ENDIF.

*        IF lv_update_batch NE zif_gbl_constants~gc_x.
    LOOP AT lt_items ASSIGNING FIELD-SYMBOL(<lfs_item>).
      ASSIGN gt_mch2[ matnr = <lfs_item>-material charg = <lfs_item>-batch ] TO FIELD-SYMBOL(<lfs_mch2>).
      IF sy-subrc IS NOT INITIAL.
        lt_mch1 = VALUE #( BASE lt_mch1
                          ( matnr = <lfs_item>-material
                            charg = <lfs_item>-batch
                            vfdat = <lfs_item>-expirydate
                            hsdat = <lfs_item>-prod_date ) ).
      ENDIF.
    ENDLOOP.

    CALL FUNCTION 'VB_UPDATE_BATCH'
      TABLES
        zmch1 = lt_mch1.
    IF sy-subrc IS INITIAL.
      "!This flag will be set when POST is clicked
      IF gv_flag_postdata = zif_gbl_constants=>gc_x.
        lv_message = |{ ls_materialdocument } { ls_materialdocyear }|.
      ELSE.
        lv_message = zif_gbl_constants=>gc_ok_to_proceed.
      ENDIF.
    ELSE.
      lv_message = ls_return-message.
    ENDIF.

    READ TABLE lt_return INTO DATA(ls_ret) INDEX 1.

*    lv_quantity = ls_items-entry_qnt.
    IF lt_return IS INITIAL.
      lv_status = zif_gbl_constants=>gc_green.
    ELSE.
      lv_status = zif_gbl_constants=>gc_red.
      IF ls_ret-message IS NOT INITIAL.    " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
        lv_message = ls_ret-message.       " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
      ENDIF.                               " R6-ENDO/FTS.EXT.069/VMANEM/EC2K902748
    ENDIF.

    IF ls_ret-type EQ zif_gbl_constants=>gc_e.
      gv_error = abap_true.
    ENDIF.

    lv_mtsnr = lwa_header-ref_doc_no_long . "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    LOOP AT lt_items ASSIGNING FIELD-SYMBOL(<lfs_item1>).

      lv_quantity = <lfs_item1>-entry_qnt.
      LOOP AT lt_serial ASSIGNING FIELD-SYMBOL(<lfs_serial>).
        IF sy-tabix NE 1.
          CLEAR: <lfs_item>, lv_quantity, lwa_header-header_txt, lv_status, lv_message .
        ENDIF.

        z_populate_alv_data(
           EXPORTING
             i_v_status   = lv_status
             i_v_header   = lwa_header-header_txt
             i_v_plant    = <lfs_item1>-plant
             i_v_sloc     = <lfs_item1>-stge_loc
             i_v_customer = <lfs_item1>-customer
             i_v_material = <lfs_item1>-material
             i_v_batch    = <lfs_item1>-batch
             i_v_quantity = lv_quantity
             i_v_unit     = <lfs_item1>-entry_uom
             i_v_serialno = <lfs_serial>-serialno
             i_v_move_type  = <lfs_item1>-move_type
             i_v_stck_type  = <lfs_item1>-stck_type
             i_v_spec_stock = <lfs_item1>-spec_stock
             i_v_prod_date  = <lfs_item1>-prod_date
             i_v_expirydate = <lfs_item1>-expirydate
             i_v_grund      = <lfs_item1>-move_reas  " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
             i_v_mat_slip   = lv_mtsnr "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
             i_v_message  = lv_message ).
      ENDLOOP.
      IF sy-subrc NE 0.
        z_populate_alv_data(
         EXPORTING
           i_v_status   = lv_status
           i_v_header   = lwa_header-header_txt
           i_v_plant    = <lfs_item1>-plant
           i_v_sloc     = <lfs_item1>-stge_loc
           i_v_customer = <lfs_item1>-customer
           i_v_material = <lfs_item1>-material
           i_v_batch    = <lfs_item1>-batch
           i_v_quantity = lv_quantity
           i_v_unit     = <lfs_item1>-entry_uom
           i_v_serialno = space
           i_v_move_type  = <lfs_item1>-move_type
           i_v_stck_type  = <lfs_item1>-stck_type
           i_v_spec_stock = <lfs_item1>-spec_stock
           i_v_prod_date  = <lfs_item1>-prod_date
           i_v_expirydate = <lfs_item1>-expirydate
           i_v_grund      = <lfs_item1>-move_reas " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
           i_v_mat_slip   = lv_mtsnr             "++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
           i_v_message  = lv_message ).
      ENDIF.
    ENDLOOP.

    CLEAR: lt_items, lt_serial.
  ENDMETHOD.
**********************************************************************
* Method Description :Methods to set the file path private variable  *
* -->  iv_file_path TYPE RLGRAP-FILENAME                             *
**********************************************************************
  METHOD z_set_file_path.

    lv_file_name = i_file_path.

  ENDMETHOD.
**********************************************************************
* Method Description : Constructor Method to get File URL            *
**********************************************************************
  METHOD constructor.
    lv_file_name = i_filepath.
  ENDMETHOD.


* Begin of Change: R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
* Method z_create_template_header_table has bene replaced by another method
**********************************************************************
*  Method to create Internal table which will be used to create Excel*
*  Template                                                          *
**********************************************************************
  METHOD z_create_template_header_table.
*
*    "File Excel data
*    lt_itab = VALUE #(  (
*                                  col1 =   TEXT-004 "document header text
*                                  col2 =   TEXT-005 "plant
*                                  col3 =   TEXT-006 "storage location
*                                  col4 =   TEXT-015 "customer
*                                  col5 =   TEXT-007 "material
*                                  col6 =   TEXT-008 "batch
*                                  col7 =   TEXT-009 "quantity
*                                  col8 =   TEXT-010 "unit of entry
*                                  col9 =   TEXT-011 "serial number
*                                  col10 =  TEXT-012 "movement type: double check
*                                  col11 =  TEXT-013 "stock type
*                                  col12 =  TEXT-014 "special stock indicator
*                                  col13 =  TEXT-026 "DOM
*                                  col14 =  TEXT-027 "SLED
*                                  col15 =  TEXT-040 "Reason Code " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
*                                  ) ).
*    lt_itab = VALUE #(  BASE  lt_itab (
*                                  col1 =   TEXT-018 "optional
*                                  col2 =   TEXT-017 "mandatory
** Begin of Changes for R6-ENDO/BUG 669327/VMANEM/EC1K916745
**                                 col3 =   TEXT-017 "mandatory
*                                  col3  = COND #( WHEN p_rad9  IS NOT INITIAL THEN TEXT-018
*                                                  WHEN p_rad10 IS NOT INITIAL THEN TEXT-018
*                                                  ELSE TEXT-017 )
**                                 col4 =   TEXT-017 "mandatory
*                                  col4 =  COND #( WHEN p_rad9  IS NOT INITIAL THEN TEXT-017
*                                                  WHEN p_rad10 IS NOT INITIAL THEN TEXT-017
*                                                  ELSE TEXT-018 )
** End of Changes for R6-ENDO/BUG 669327/VMANEM/EC1K916745
*                                  col5 =   TEXT-017 "mandatory
*                                  col6 =   TEXT-019 "mandatory if batch is managed
*                                  col7 =   TEXT-017 "mandatory
*                                  col8 =   TEXT-017 "mandatory
*                                  col9 =   TEXT-020 "Mandatory if Material is Serial Managed
*                                  col10 =  TEXT-017 "mandatory
*                                  col11 =  TEXT-017 "mandatory
*                                  col12 =  TEXT-021 "Only for option 5 & 6
*                                  col13 =  TEXT-018 "optional
*                                  col14 =  TEXT-018 "optional
*                                  col15 =  TEXT-018 "optional " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
*                                  ) ).
  ENDMETHOD.

**********************************************************************
*  R7: Method to create Dynamic Internal table based upon the        *
*   Movement Type which will be used to create Excel Template        *
**********************************************************************
  METHOD z_create_dyn_template_hdr_tab.


* Dynamic Structure for Selecged Movement Types
    TYPES :
      BEGIN OF lty_dyn_fld, "type for final internal table
        fld_nm(40) TYPE c,
      END OF  lty_dyn_fld.

    DATA: lt_dyn_fld  TYPE STANDARD TABLE OF lty_dyn_fld,
          lwa_dyn_fld LIKE LINE OF lt_dyn_fld.

*Field Catalog data
    DATA: lt_dyn_fcat  TYPE  lvc_t_fcat,
          lwa_fieldcat TYPE lvc_s_fcat.

    FIELD-SYMBOLS:<lfs_dyn_table> TYPE STANDARD TABLE,
                  <lfs_dyn_wa>    TYPE any.
    FIELD-SYMBOLS: <lfs_field_to> TYPE any.
    DATA: lv_fieldname TYPE lvc_fname.
    DATA: lwa_itab LIKE LINE OF lt_itab.

    DATA: lt_dy_table TYPE REF TO data,
          lwa_dy_line TYPE REF TO data.


    CASE: p_mtyp.
      WHEN zif_global_constants_n3=>gc_movement_711   " '711'
         OR zif_global_constants_n3=>gc_movement_712. " '712'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
        lwa_dyn_fld-fld_nm = TEXT-006. APPEND lwa_dyn_fld TO lt_dyn_fld. "Storage Location
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type
        lwa_dyn_fld-fld_nm = TEXT-013. APPEND lwa_dyn_fld TO lt_dyn_fld. "Stock Type
*        lwa_dyn_fld-fld_nm = TEXT-040. APPEND lwa_dyn_fld TO lt_dyn_fld. "Reason Code

      WHEN zif_global_constants_n3=>gc_movement_711w   " '711W'
         OR zif_global_constants_n3=>gc_movement_712w. " '712W'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
        lwa_dyn_fld-fld_nm = TEXT-015. APPEND lwa_dyn_fld TO lt_dyn_fld. "Customer
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-058. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material Slip
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type
        lwa_dyn_fld-fld_nm = TEXT-013. APPEND lwa_dyn_fld TO lt_dyn_fld. "Stock Type
        lwa_dyn_fld-fld_nm = TEXT-014. APPEND lwa_dyn_fld TO lt_dyn_fld. "Special Stock Indicator
        lwa_dyn_fld-fld_nm = TEXT-040. APPEND lwa_dyn_fld TO lt_dyn_fld. "Reason Code

      WHEN zif_gbl_constants=>gc_movement_971 OR zif_gbl_constants=>gc_movement_972.    " '971' OR '972'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
        lwa_dyn_fld-fld_nm = TEXT-006. APPEND lwa_dyn_fld TO lt_dyn_fld. "Storage Location
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type
        lwa_dyn_fld-fld_nm = TEXT-013. APPEND lwa_dyn_fld TO lt_dyn_fld. "Stock Type
        lwa_dyn_fld-fld_nm = TEXT-026. APPEND lwa_dyn_fld TO lt_dyn_fld. "DOM
        lwa_dyn_fld-fld_nm = TEXT-027. APPEND lwa_dyn_fld TO lt_dyn_fld. "SLED

      WHEN zif_global_constants_n3=>gc_movement_971w
        OR zif_global_constants_n3=>gc_movement_972w. " '971W' OR '972W'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
        lwa_dyn_fld-fld_nm = TEXT-015. APPEND lwa_dyn_fld TO lt_dyn_fld. "Customer
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type
        lwa_dyn_fld-fld_nm = TEXT-013. APPEND lwa_dyn_fld TO lt_dyn_fld. "Stock Type
        lwa_dyn_fld-fld_nm = TEXT-014. APPEND lwa_dyn_fld TO lt_dyn_fld. "Special Stock Indicator
        lwa_dyn_fld-fld_nm = TEXT-026. APPEND lwa_dyn_fld TO lt_dyn_fld. "DOM
        lwa_dyn_fld-fld_nm = TEXT-027. APPEND lwa_dyn_fld TO lt_dyn_fld. "SLED

      WHEN zif_gbl_constants=>gc_movement_997 OR zif_gbl_constants=>gc_movement_998.  " '997' OR '998'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
        lwa_dyn_fld-fld_nm = TEXT-006. APPEND lwa_dyn_fld TO lt_dyn_fld. "Storage Location
        lwa_dyn_fld-fld_nm = TEXT-015. APPEND lwa_dyn_fld TO lt_dyn_fld. "Customer
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type

      WHEN zif_global_constants_n3=>gc_movement_551w
        OR zif_global_constants_n3=>gc_movement_552w.  " '551W' OR '552W'.
        lwa_dyn_fld-fld_nm = TEXT-004. APPEND lwa_dyn_fld TO lt_dyn_fld. "Document Header Text
        lwa_dyn_fld-fld_nm = TEXT-005. APPEND lwa_dyn_fld TO lt_dyn_fld. "Plant
*        lwa_dyn_fld-fld_nm = TEXT-006. APPEND lwa_dyn_fld TO lt_dyn_fld. "Storage Location
        lwa_dyn_fld-fld_nm = TEXT-015. APPEND lwa_dyn_fld TO lt_dyn_fld. "Customer
        lwa_dyn_fld-fld_nm = TEXT-007. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material
        lwa_dyn_fld-fld_nm = TEXT-058. APPEND lwa_dyn_fld TO lt_dyn_fld. "Material Slip
        lwa_dyn_fld-fld_nm = TEXT-008. APPEND lwa_dyn_fld TO lt_dyn_fld. "Batch
        lwa_dyn_fld-fld_nm = TEXT-009. APPEND lwa_dyn_fld TO lt_dyn_fld. "Quantity
        lwa_dyn_fld-fld_nm = TEXT-010. APPEND lwa_dyn_fld TO lt_dyn_fld. "Unit of Entry
        lwa_dyn_fld-fld_nm = TEXT-011. APPEND lwa_dyn_fld TO lt_dyn_fld. "Serial Number
        lwa_dyn_fld-fld_nm = TEXT-012. APPEND lwa_dyn_fld TO lt_dyn_fld. "Movement type
        lwa_dyn_fld-fld_nm = TEXT-013. APPEND lwa_dyn_fld TO lt_dyn_fld. "Stock Type
        lwa_dyn_fld-fld_nm = TEXT-014. APPEND lwa_dyn_fld TO lt_dyn_fld. "Special Stock Indicator
        lwa_dyn_fld-fld_nm = TEXT-040. APPEND lwa_dyn_fld TO lt_dyn_fld. "Reason Code

    ENDCASE.

    " Dynamic Header Field Range
    gv_dynl = lines( lt_dyn_fld ).
    gv_dynl = gv_dynl - 1.

    DATA(lv_incr) = 1.
    "Get fields
    LOOP AT lt_dyn_fld  INTO lwa_dyn_fld.
      CLEAR lwa_fieldcat.
      lwa_fieldcat-fieldname  = TEXT-077 && lv_incr.
      lwa_fieldcat-ifieldname = lwa_dyn_fld-fld_nm.
      lwa_fieldcat-datatype   = TEXT-078.  " 'CHAR'.
      lwa_fieldcat-inttype    = TEXT-079.  " 'C'.
      lwa_fieldcat-intlen     = '40'.
*      gwa_fieldcat-decimals  = lwa_tabdescr-decimals.
      lv_incr = lv_incr + 1.
      APPEND lwa_fieldcat TO lt_dyn_fcat.
    ENDLOOP.
    "Create Dynamic Table Structure
    CALL METHOD cl_alv_table_create=>create_dynamic_table
      EXPORTING
        it_fieldcatalog           = lt_dyn_fcat
      IMPORTING
        ep_table                  = lt_dy_table
      EXCEPTIONS
        generate_subpool_dir_full = 1
        OTHERS                    = 2.

    IF sy-subrc EQ 0.
      ASSIGN lt_dy_table->* TO <lfs_dyn_table>.
    ENDIF.

* Create Dynamic Work Area and Assign to FS
    CREATE DATA lwa_dy_line LIKE LINE OF <lfs_dyn_table>.
    ASSIGN lwa_dy_line->* TO <lfs_dyn_wa>.

    LOOP AT lt_dyn_fld INTO lwa_dyn_fld.
      lv_fieldname =  TEXT-077 && sy-tabix.
      ASSIGN COMPONENT lv_fieldname OF STRUCTURE <lfs_dyn_wa> TO <lfs_field_to>.
      <lfs_field_to> = lwa_dyn_fld-fld_nm.
    ENDLOOP.

    IF <lfs_dyn_wa> IS ASSIGNED.
      APPEND  <lfs_dyn_wa> TO <lfs_dyn_table>.
    ENDIF.

    CLEAR: lt_itab.
    LOOP AT <lfs_dyn_table> INTO <lfs_dyn_wa>.
      MOVE-CORRESPONDING <lfs_dyn_wa> TO lwa_itab.
      APPEND lwa_itab TO lt_itab.
    ENDLOOP.

  ENDMETHOD.
* End of Change: R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

**********************************************************************
* Method Description : Method to show ALV Output                     *
**********************************************************************
  METHOD z_display_alv.
    "local constant
    CONSTANTS: lc_quantity TYPE lvc_fname VALUE zif_gbl_constants=>gc_quantity1,
               lc_status   TYPE lvc_fname VALUE zif_gbl_constants=>gc_status,
               lc_uom      TYPE lvc_fname VALUE 'UNIT',
               lc_ground   TYPE lvc_fname VALUE 'GRUND'.  " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    "local variable
    DATA : lv_quantity TYPE scrtext_s,
           lv_status   TYPE scrtext_s,
           lv_qty_m    TYPE scrtext_m,  " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
           lv_qty_l    TYPE scrtext_l.   " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

*    lv_quantity = lc_quantity.   " -- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lv_quantity = TEXT-009.       " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lv_qty_m    = TEXT-009.       " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lv_qty_l    = TEXT-009.       " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lv_status = lc_status.

    TRY.
        CALL METHOD cl_salv_table=>factory
          IMPORTING
            r_salv_table = DATA(lo_alv)   " Basis Class Simple ALV Tables
          CHANGING
            t_table      = lt_alv_output.
      CATCH cx_salv_msg INTO DATA(lo_salv_msg).
        WRITE : lo_salv_msg->get_text( ).
    ENDTRY.

    DATA(lo_columns) = lo_alv->get_columns( ).
    DATA(lo_column)  = lo_columns->get_column( lc_quantity ).
    lo_column->set_zero( abap_false ).

    DATA(lo_disp_set) = lo_alv->get_display_settings( ).
    lo_disp_set->set_list_header( TEXT-025 ). " ALV Output
    lo_disp_set->set_striped_pattern( value =  cl_salv_display_settings=>true ).

    DATA(lo_salv_col) = lo_columns->get_column( lc_quantity ).
    lo_salv_col->set_short_text( lv_quantity ).
    lo_salv_col->set_medium_text( lv_qty_m ).                    " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lo_salv_col->set_long_text( lv_qty_l ).                      " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lo_salv_col->set_alignment( if_salv_c_alignment=>right ).    " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    lo_salv_col = lo_columns->get_column( lc_uom ).              " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lo_salv_col->set_alignment( if_salv_c_alignment=>right ).    " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

    lo_salv_col = lo_columns->get_column( lc_status ).
    lo_salv_col->set_short_text( lv_status ).

    lo_salv_col = lo_columns->get_column( lc_ground ).          " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
    lo_salv_col->set_leading_zero( abap_true ).                 " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625


*    ALV-Toolbar
    IF gv_flag_setstatus = ' '.
      lo_alv->set_screen_status(
        pfstatus      = zif_gbl_constants=>gc_standard
        report        = sy-repid
        set_functions = lo_alv->c_functions_all ).
    ELSE.
      lo_alv->set_screen_status(
          pfstatus      = zif_gbl_constants=>gc_standard1
          report        = sy-repid
          set_functions = lo_alv->c_functions_all ).
    ENDIF.
    TRY.
        go_events_alv = lo_alv->get_event( ).
      CATCH cx_salv_not_found INTO DATA(lv_salv_nf).
        MESSAGE lv_salv_nf->get_text( ) TYPE zif_gbl_constants=>gc_e.
    ENDTRY.

    CREATE OBJECT go_events
      EXPORTING
        i_filepath = lv_file_name.

    SET HANDLER go_events->z_on_user_command FOR lo_alv->get_event( ).
    lo_alv->display( ).                       " Display final ALV

  ENDMETHOD.
**********************************************************************
* Method Description : Method to handle user command                 *
**********************************************************************
  METHOD z_on_click.

    IF sy-ucomm EQ zif_gbl_constants=>gc_proceed.

      IF gv_error EQ ' '.
        gv_flag_postdata = zif_gbl_constants=>gc_x.
        gv_flag_setstatus = zif_gbl_constants=>gc_x.
*        z_populate_bapi_tables( ).      " -- R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        z_populate_bapi_tables_r7( ).    " ++ R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
        z_display_alv( ).
      ELSE.
        MESSAGE e009.   "Please correct the data in file!
      ENDIF.

    ELSEIF sy-ucomm EQ zif_gbl_constants=>gc_exit.
      SET SCREEN 0.
    ENDIF.
  ENDMETHOD.
*********************************************************************
* Method Description : Method to insert messages in Output ALV      *
*********************************************************************
  METHOD z_populate_alv_data.

    DATA: lv_qnt TYPE char20.
    CLEAR: lv_qnt.
    IF i_v_quantity NE 0.
      lv_qnt = i_v_quantity.
    ENDIF.

    lt_alv_output = VALUE #( BASE lt_alv_output
                            ( status     = i_v_status
                              header_txt = i_v_header
                              plant      = i_v_plant
                              sloc       = i_v_sloc
                              customer   = i_v_customer
                              material   = i_v_material
                              batch      = i_v_batch
                              quantity   = lv_qnt
                              unit       = i_v_unit
                              serialno   = i_v_serialno
                              move_type  = i_v_move_type
                              stck_type  = i_v_stck_type
                              spec_stock = i_v_spec_stock
                              prod_date  = i_v_prod_date
                              expirydate = i_v_expirydate
                              grund      = i_v_grund      " R6-ENDO/FTS.EXT.069/CR1199/VMANEM/EC2K902748
                              mat_slip   = i_v_mat_slip   " R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
                              message    = i_v_message
                            ) ) .

  ENDMETHOD.
**********************************************************************
* Method Description : Method to populate serial tab for BAPI        *
**********************************************************************
  METHOD z_populate_bapi_serial_data.
    IF i_wa_data-serialno IS NOT INITIAL.
      lt_serial = VALUE #( BASE lt_serial ( matdoc_itm = i_matid
                                            serialno   = i_wa_data-serialno ) ).
    ENDIF.
  ENDMETHOD.

* Begin of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625
**********************************************************************
* Method Description : Method to set movement type for R7            *
**********************************************************************
  METHOD  z_set_movtype_r7 .
    lv_move_type   = p_mtyp(3).                 " Movementy Type
    IF p_mtyp+3(1) IS NOT INITIAL.
      lv_spec_stock  = zif_gbl_constants=>gc_x. " Special Stock Ind.
    ENDIF.
  ENDMETHOD.
* End of Changes R7-ORTHO/SDE2/FTS.EXT.286/ EC2K904625

ENDCLASS.

CLASS lcl_handle_events IMPLEMENTATION.
*-----------------------------------------------------------------------------------*
*&Method  on_user_command (To post goods movement and create material document)
*-----------------------------------------------------------------------------------*
  METHOD z_on_user_command.
    z_on_click( ).
  ENDMETHOD. " z on user command

ENDCLASS.
