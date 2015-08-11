
"-Begin-----------------------------------------------------------------
  Report zTEST.

    "-Structures--------------------------------------------------------
      Data: Begin Of TestStruc,
              ElemInt  Type i,
              ElemPack Type p,
              ElemFlt  Type f,
              ElemStr  Type String,
            End Of TestStruc.

    "-Variables---------------------------------------------------------
      Data l_str_Test Like TestStruc.
      Data l_tab_Test Like Standard Table Of TestStruc.
      Data l_rcl_excel Type Ref To ZCL_EXCEL.
      Data l_rcl_worksheet Type Ref To ZCL_EXCEL_WORKSHEETS.
      Data l_rif_excel_writer Type Ref To ZIF_EXCEL_WRITER.
      Data l_cnt_Lines Type i Value 0.
      Data l_xlsx_datastream Type XString.
      Data l_dTab Type Table Of x255.
      Data l_len Type i Value 0.
      Data l_t0 Type i Value 0.
      Data l_t1 Type i Value 0.

    "-Main--------------------------------------------------------------
      Do 16384 Times.
        l_str_Test-ElemInt = sy-index.
        l_str_Test-ElemPack = sy-index * 1000.
        l_str_Test-ElemFlt = sy-index * '3.14'.
        l_str_Test-ElemStr = `Dies ist ein Test ` && sy-index.
        Append l_str_Test To l_tab_Test.
      EndDo.

      If l_tab_Test Is Not Initial.
        Describe Table l_tab_Test Lines l_cnt_Lines.
        If l_cnt_Lines <= 1048576.
          Create Object l_rcl_excel.
          l_rcl_worksheet = l_rcl_excel->get_active_worksheet( ).
          l_rcl_worksheet->bind_table( ip_table = l_tab_Test ).

          Get Run Time Field l_t0.

          Create Object l_rif_excel_writer Type ZCL_EXCEL_WRITER_2007.
          "Create Object l_rif_excel_writer Type ZCL_EXCEL_WRITER_HUGE_FILE.

          l_xlsx_datastream = l_rif_excel_writer->write_file( l_rcl_excel ).

          Get Run Time Field l_t1.

          l_t0 = l_t1 - l_t0.
          Write: l_t0.

          Call Function 'SCMS_XSTRING_TO_BINARY'
            Exporting
              BUFFER = l_xlsx_datastream
            Importing
              OUTPUT_LENGTH = l_len
            Tables
              BINARY_TAB = l_dTab.

          Call Function 'GUI_DOWNLOAD'
            Exporting
              BIN_FILESIZE = l_len
              FILENAME = 'C:\Dummy\Test.xlsx'
              FILETYPE = 'BIN'
            Tables
              DATA_TAB = l_dTab
            Exceptions
              Others = 1.

        EndIf.
      EndIf.

"-End-------------------------------------------------------------------
