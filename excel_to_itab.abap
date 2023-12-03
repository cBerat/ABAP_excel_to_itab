*&---------------------------------------------------------------------*
*& Report ZBC_EXCEL_TO_ITAB
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT ZBC_EXCEL_TO_ITAB.

TYPES: BEGIN OF gty_itab,
         col1 type char20,
         col2 type char20,
         col3 type char20,
         col4 type char20,
         col5 type char20,
       END OF gty_itab.


DATA: gt_intern TYPE TABLE OF ALSMEX_TABLINE,
      gs_intern TYPE ALSMEX_TABLINE,
      gt_itab   TYPE TABLE OF gty_itab,
      gs_itab   TYPE gty_itab.

FIELD-SYMBOLS: <gfv_value> TYPE any.

PARAMETERS: p_file type rlgrap-FILENAME.


at SELECTION-SCREEN on VALUE-REQUEST FOR p_file.

  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
*     PROGRAM_NAME        = SYST-CPROG
*     DYNPRO_NUMBER       = SYST-DYNNR
      FIELD_NAME = 'p_file'
    IMPORTING
      FILE_NAME  = p_file.

START-OF-SELECTION.

  CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      FILENAME                = p_file
      I_BEGIN_COL             = 2
      I_BEGIN_ROW             = 2
      I_END_COL               = 4
      I_END_ROW               = 4
    TABLES
      INTERN                  = gt_intern
    EXCEPTIONS
      INCONSISTENT_PARAMETERS = 1
      UPLOAD_OLE              = 2
      OTHERS                  = 3.
  IF SY-SUBRC <> 0.
* Implement suitable error handling here
  ENDIF.

  LOOP AT gt_intern INTO gs_intern.
    ASSIGN COMPONENT gs_intern-COL of STRUCTURE gs_itab to <gfv_value>.
    <gfv_value> = GS_INTERN-VALUE.
    at END OF row. " row her degistiginde burasý calýsacak.
      append gs_itab to gt_itab.
    endat.
  ENDLOOP.
