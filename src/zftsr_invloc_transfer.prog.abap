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
***************************************************************************
REPORT zftsr_invloc_transfer MESSAGE-ID zptp NO STANDARD PAGE HEADING.

CLASS lcl_invloc_transfer DEFINITION DEFERRED.

* Include program for top declaration
INCLUDE zftsr_invloc_transfer_top.

* Include for selection screen
INCLUDE zftsr_invloc_transfer_f02.  " ++ R6-ENDO/BUG 669327/VMANEM/EC1K916745

* Include program for class definition and implementation
INCLUDE zftsr_invloc_transfer_f01.

** Include for selection screen
*INCLUDE zftsr_invloc_transfer_f02.  " -- R6-ENDO/BUG 669327/VMANEM/EC1K916745

*Include for selection screen events
INCLUDE zftsr_invloc_transfer_f03.
