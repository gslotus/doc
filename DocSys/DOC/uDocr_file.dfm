object fmDocr_file: TfmDocr_file
  Left = 48
  Top = 62
  Width = 1016
  Height = 609
  Caption = #35373#23450#22577#34920#26597#30475#20154#21729
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label7: TLabel
    Left = 88
    Top = 40
    Width = 46
    Height = 13
    Caption = 'lab_dst01'
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 281
    Height = 571
    Align = alLeft
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 9
      Top = 112
      Width = 36
      Height = 18
      Caption = #37096#38272
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 8
      Top = 168
      Width = 108
      Height = 18
      Caption = #26410#35373#27402#38480#21517#21934
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 8
      Top = 8
      Width = 72
      Height = 18
      Caption = #22823#39006#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 160
      Top = 8
      Width = 72
      Height = 18
      Caption = #23567#39006#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 8
      Top = 32
      Width = 36
      Height = 18
      Caption = #24207#34399
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lab_dst01: TLabel
      Left = 88
      Top = 8
      Width = 5
      Height = 18
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lab_dst02: TLabel
      Left = 240
      Top = 8
      Width = 5
      Height = 18
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lab_dst03: TLabel
      Left = 48
      Top = 32
      Width = 5
      Height = 18
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label8: TLabel
      Left = 8
      Top = 136
      Width = 36
      Height = 18
      Caption = #22995#21517
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 8
      Top = 56
      Width = 72
      Height = 18
      Caption = #22577#34920#21517#31281
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object lab_repdesc: TLabel
      Left = 88
      Top = 56
      Width = 5
      Height = 18
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label10: TLabel
      Left = 9
      Top = 85
      Width = 36
      Height = 18
      Caption = #24288#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object cb_dep: TComboBox
      Left = 48
      Top = 109
      Width = 185
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemHeight = 18
      ParentFont = False
      TabOrder = 0
    end
    object DBGrid1: TDBGrid
      Left = 1
      Top = 202
      Width = 279
      Height = 368
      Align = alBottom
      DataSource = DataSource2
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      TitleFont.Charset = ANSI_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -18
      TitleFont.Name = #26032#32048#26126#39636
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'GEN01'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'GEN02'
          Visible = True
        end>
    end
    object btn_all: TButton
      Left = 136
      Top = 168
      Width = 137
      Height = 25
      Caption = #35373#23450#20840#39636#20154#21729
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btn_allClick
    end
    object Button1: TButton
      Left = 176
      Top = 136
      Width = 73
      Height = 25
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = Button1Click
    end
    object edt_name: TEdit
      Left = 48
      Top = 136
      Width = 121
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
    end
    object cb_fac: TComboBox
      Left = 47
      Top = 82
      Width = 185
      Height = 26
      DropDownCount = 10
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemHeight = 18
      ParentFont = False
      TabOrder = 5
      OnChange = cb_facChange
    end
  end
  object Panel2: TPanel
    Left = 281
    Top = 0
    Width = 719
    Height = 571
    Align = alClient
    TabOrder = 1
    object Panel3: TPanel
      Left = 41
      Top = 1
      Width = 677
      Height = 569
      Align = alClient
      Color = clMoneyGreen
      TabOrder = 0
      object Label3: TLabel
        Left = 8
        Top = 8
        Width = 108
        Height = 18
        Caption = #24050#35373#27402#38480#21517#21934
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -18
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label11: TLabel
        Left = 11
        Top = 39
        Width = 77
        Height = 18
        Caption = #39023#31034#24288#21029':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -18
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object DBGrid2: TDBGrid
        Left = 1
        Top = 71
        Width = 675
        Height = 497
        Align = alBottom
        DataSource = DataSource1
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -18
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
        TitleFont.Charset = ANSI_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -18
        TitleFont.Name = #26032#32048#26126#39636
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'DOCR06'
            Title.Caption = #24288#21029
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR04'
            Width = 79
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'GEN02'
            Width = 77
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR05'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR01'
            Visible = False
          end
          item
            Expanded = False
            FieldName = 'DOCR02'
            Visible = False
          end
          item
            Expanded = False
            FieldName = 'DOCR03'
            Visible = False
          end
          item
            Expanded = False
            FieldName = 'DOCR_ACTI'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR_USER'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR_DATE'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DOCR_FAC'
            Visible = False
          end>
      end
      object cb_fac1: TComboBox
        Left = 92
        Top = 35
        Width = 185
        Height = 26
        DropDownCount = 10
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -18
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ItemHeight = 18
        ParentFont = False
        TabOrder = 1
        OnChange = cb_fac1Change
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 40
      Height = 569
      Align = alLeft
      Color = clMoneyGreen
      TabOrder = 1
      object sb_remove: TSpeedButton
        Left = 0
        Top = 280
        Width = 41
        Height = 33
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333FF3333333333333003333333333333F77F33333333333009033
          333333333F7737F333333333009990333333333F773337FFFFFF330099999000
          00003F773333377777770099999999999990773FF33333FFFFF7330099999000
          000033773FF33777777733330099903333333333773FF7F33333333333009033
          33333333337737F3333333333333003333333333333377333333333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333333333333}
        NumGlyphs = 2
        OnClick = sb_removeClick
      end
      object sb_add: TSpeedButton
        Left = 0
        Top = 240
        Width = 41
        Height = 33
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333333333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333FF3333333333333003333
          3333333333773FF3333333333309003333333333337F773FF333333333099900
          33333FFFFF7F33773FF30000000999990033777777733333773F099999999999
          99007FFFFFFF33333F7700000009999900337777777F333F7733333333099900
          33333333337F3F77333333333309003333333333337F77333333333333003333
          3333333333773333333333333333333333333333333333333333333333333333
          3333333333333333333333333333333333333333333333333333}
        NumGlyphs = 2
        OnClick = sb_addClick
      end
    end
  end
  object DOCR_FILE: TADOQuery
    Connection = dm.glm
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM DOCR_FILE')
    Left = 504
    Top = 184
    object DOCR_FILEDOCR01: TWideStringField
      DisplayLabel = #22823#39006#22823#30908
      FieldName = 'DOCR01'
      Size = 10
    end
    object DOCR_FILEDOCR02: TWideStringField
      DisplayLabel = #23567#39006#20195#30908
      FieldName = 'DOCR02'
      Size = 10
    end
    object DOCR_FILEDOCR03: TWideStringField
      DisplayLabel = #25991#20214#24207#34399
      FieldName = 'DOCR03'
      Size = 10
    end
    object DOCR_FILEDOCR04: TWideStringField
      DisplayLabel = #20154#21729#24115#34399
      DisplayWidth = 10
      FieldName = 'DOCR04'
      Size = 30
    end
    object DOCR_FILEDOCR05: TWideStringField
      DisplayLabel = 'mail address'
      DisplayWidth = 15
      FieldName = 'DOCR05'
      Size = 30
    end
    object DOCR_FILEDOCR_USER: TWideStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'DOCR_USER'
      Size = 10
    end
    object DOCR_FILEDOCR_DATE: TDateTimeField
      DisplayLabel = #30064#21205#26085#26399
      DisplayWidth = 10
      FieldName = 'DOCR_DATE'
    end
    object DOCR_FILEDOCR_ACTI: TWideStringField
      DisplayLabel = #26159#21542#21487#29992
      FieldName = 'DOCR_ACTI'
      Size = 1
    end
    object DOCR_FILEDOCR_FAC: TWideStringField
      DisplayLabel = #24288#21029
      FieldName = 'DOCR_FAC'
      Size = 10
    end
    object DOCR_FILEGEN02: TStringField
      DisplayLabel = #22995#21517
      FieldName = 'GEN02'
    end
    object DOCR_FILEDOCR06: TWideStringField
      FieldName = 'DOCR06'
      Size = 10
    end
  end
  object DataSource1: TDataSource
    DataSet = DOCR_FILE
    Left = 432
    Top = 144
  end
  object docr1_file: TADOQuery
    Connection = dm.glm
    Parameters = <>
    Left = 168
    Top = 272
    object docr1_fileGEN01: TStringField
      DisplayLabel = #24115#34399
      FieldName = 'GEN01'
      Size = 10
    end
    object docr1_fileGEN02: TStringField
      DisplayLabel = #22995#21517
      FieldName = 'GEN02'
    end
  end
  object DataSource2: TDataSource
    DataSet = docr1_file
    Left = 72
    Top = 264
  end
end
