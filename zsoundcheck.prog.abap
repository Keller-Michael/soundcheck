*&---------------------------------------------------------------------*
*& Report zsoundcheck
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zsoundcheck.

SELECTION-SCREEN: BEGIN OF BLOCK bl_control WITH FRAME TITLE gv_tx001,
                  BEGIN OF LINE,
                  PUSHBUTTON 1(20)  pb_play USER-COMMAND play VISIBLE LENGTH 10,
                  PUSHBUTTON 13(20) pb_stop USER-COMMAND stop VISIBLE LENGTH 10,
                  END OF LINE,
                  END OF BLOCK bl_control.

**********************************************************************

CLASS lcx_wmp_error DEFINITION INHERITING FROM cx_static_check.

  PUBLIC SECTION.
    INTERFACES if_t100_dyn_msg.

    DATA: mv_subrc LIKE sy-subrc,
          mv_msgli LIKE sy-msgli.

    METHODS constructor
      IMPORTING
        iv_subrc LIKE sy-subrc OPTIONAL
        iv_msgli LIKE sy-msgli OPTIONAL.

ENDCLASS.

CLASS lcx_wmp_error IMPLEMENTATION.

  METHOD constructor.
    super->constructor( ).

    " compare SAP Online Help, sy-subrc and sy-msgli are important when using OLE
    mv_subrc = iv_subrc.
    mv_msgli = iv_msgli.
  ENDMETHOD.
ENDCLASS.

**********************************************************************

CLASS lcl_wmp DEFINITION.

  PUBLIC SECTION.
    METHODS constructor
      IMPORTING
        iv_url       TYPE string OPTIONAL
        iv_autostart TYPE abap_bool DEFAULT abap_false
      RAISING
        lcx_wmp_error.

    METHODS play
      RAISING
        lcx_wmp_error.

    METHODS stop
      RAISING
        lcx_wmp_error.

    METHODS set_autostart_on
      RAISING
        lcx_wmp_error.

    METHODS set_loop_mode
      RAISING
        lcx_wmp_error.

    METHODS set_autostart_off
      RAISING
        lcx_wmp_error.

  PRIVATE SECTION.
    DATA: mv_player   TYPE ole2_object,
          mv_control  TYPE ole2_object,
          mv_settings TYPE ole2_object.

ENDCLASS.

**********************************************************************

CLASS lcl_wmp IMPLEMENTATION.

  METHOD constructor.
    CREATE OBJECT mv_player 'WMP'.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'No connection to Windows Media player.'.
    ENDIF.

    GET PROPERTY OF mv_player 'settings' = mv_settings.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'No settings.'.
    ENDIF.

    GET PROPERTY OF mv_player 'controls' = mv_control.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'No control.'.
    ENDIF.

    IF iv_autostart = abap_false.
      me->set_autostart_off( ).
    ELSE.
      me->set_autostart_on( ).
    ENDIF.

    IF iv_url IS NOT INITIAL.
      SET PROPERTY OF mv_player 'URL' = iv_url.
      IF sy-subrc <> 0.
        RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'URL not set.'.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD play.
    CALL METHOD OF mv_control 'play'.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'Not played due to error.'.
    ELSE.
      MESSAGE 'Windows Media Player is playing.' TYPE 'S'.
    ENDIF.
  ENDMETHOD.

  METHOD stop.
    CALL METHOD OF mv_control 'stop'.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'Not stopped due to error.'.
    ELSE.
      MESSAGE 'Windows Media Player stopped playing.' TYPE 'S'.
    ENDIF.
  ENDMETHOD.

  METHOD set_autostart_off.
    SET PROPERTY OF mv_settings 'autoStart' = 'false'.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'Autostart not set.'.
    ENDIF.
  ENDMETHOD.

  METHOD set_autostart_on.
    SET PROPERTY OF mv_settings 'autoStart' = 'true'.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'Autostart not set.'.
    ENDIF.
  ENDMETHOD.

  METHOD set_loop_mode.
    CALL METHOD OF mv_settings 'setMode'
      EXPORTING
        #1 = 'loop'
        #2 = '1'.

    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE lcx_wmp_error MESSAGE e499(sy) WITH 'Loop mode not set.'.
    ENDIF.
  ENDMETHOD.

ENDCLASS.

**********************************************************************

TABLES sscrfields.

DATA gr_wmp TYPE REF TO lcl_wmp.

**********************************************************************

INITIALIZATION.
  PERFORM initialize.

AT SELECTION-SCREEN.
  PERFORM handle_user_input.

START-OF-SELECTION.

FORM initialize.

  gv_tx001 = 'Control Panel'.

  CALL FUNCTION 'ICON_CREATE'
    EXPORTING
      name                  = icon_checked
      text                  = 'Play'
      info                  = 'Play'
*     ADD_STDINF            = 'X'
    IMPORTING
      result                = pb_play
    EXCEPTIONS
      icon_not_found        = 1
      outputfield_too_short = 2
      OTHERS                = 3.

  IF sy-subrc <> 0.
  ENDIF.

  CALL FUNCTION 'ICON_CREATE'
    EXPORTING
      name                  = icon_incomplete
      text                  = 'Stop'
      info                  = 'Stop'
*     ADD_STDINF            = 'X'
    IMPORTING
      result                = pb_stop
    EXCEPTIONS
      icon_not_found        = 1
      outputfield_too_short = 2
      OTHERS                = 3.

  IF sy-subrc <> 0.
  ENDIF.

ENDFORM.

FORM handle_user_input.

  DATA lr_wmp_error TYPE REF TO lcx_wmp_error.

  IF sy-dynnr <> '1000'.
    RETURN.
  ENDIF.

  TRY.
      IF gr_wmp IS NOT BOUND.
        CREATE OBJECT gr_wmp
          EXPORTING
            iv_url       = 'C:\windows\media\town.mid'
            iv_autostart = abap_false.

        gr_wmp->set_loop_mode( ).
      ENDIF.

      CASE sscrfields.
        WHEN 'PLAY'.
          gr_wmp->play( ).
        WHEN 'STOP'.
          gr_wmp->stop( ).
      ENDCASE.

    CATCH lcx_wmp_error INTO lr_wmp_error.
      MESSAGE lr_wmp_error->get_text( ) TYPE 'I'.
      RETURN.
  ENDTRY.

ENDFORM.
