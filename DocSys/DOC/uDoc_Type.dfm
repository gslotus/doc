object fmdoc_type: Tfmdoc_type
  Left = 119
  Top = 91
  Width = 908
  Height = 584
  Caption = #25991#20214#39006#21029#35373#23450'(doc_type)'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 106
  TextHeight = 13
  object Panel2: TPanel
    Left = 0
    Top = 216
    Width = 892
    Height = 328
    Align = alBottom
    TabOrder = 0
    object Panel5: TPanel
      Left = 1
      Top = 1
      Width = 890
      Height = 326
      Align = alClient
      TabOrder = 0
      object DBGrid1: TDBGrid
        Left = 1
        Top = 1
        Width = 888
        Height = 324
        Align = alClient
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
            FieldName = 'DTYPE01'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE02'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE_ACTI'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE011'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE021'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE03'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE_USER'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DTYPE_DATE'
            Visible = True
          end>
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 892
    Height = 57
    Align = alTop
    Color = clMoneyGreen
    TabOrder = 1
    object Label7: TLabel
      Left = 8
      Top = 16
      Width = 72
      Height = 18
      Caption = #25215#36774#37096#38272
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object cb_dep1: TComboBox
      Left = 96
      Top = 16
      Width = 145
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
    object btn_que: TButton
      Left = 256
      Top = 16
      Width = 75
      Height = 25
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = btn_queClick
    end
    object Button1: TButton
      Left = 480
      Top = 16
      Width = 75
      Height = 25
      Caption = #26032#22686
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 555
      Top = 16
      Width = 75
      Height = 25
      Caption = #20462#25913
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 630
      Top = 16
      Width = 75
      Height = 25
      Caption = #20786#23384
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 705
      Top = 16
      Width = 75
      Height = 25
      Caption = #20316#24290
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = Button4Click
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 57
    Width = 892
    Height = 159
    Align = alClient
    Color = clSkyBlue
    TabOrder = 2
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 72
      Height = 18
      Caption = #22823#39006#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 8
      Top = 48
      Width = 72
      Height = 18
      Caption = #23567#39006#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 232
      Top = 16
      Width = 108
      Height = 18
      Caption = #22823#39006#21517#31281#35498#26126
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 232
      Top = 48
      Width = 108
      Height = 18
      Caption = #23567#39006#21517#31281#35498#26126
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 8
      Top = 88
      Width = 72
      Height = 18
      Caption = #25215#36774#37096#38272
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object edt_type01: TEdit
      Left = 96
      Top = 16
      Width = 121
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      OnExit = edt_type01Exit
    end
    object cb_dep: TComboBox
      Left = 96
      Top = 88
      Width = 145
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemHeight = 18
      ParentFont = False
      TabOrder = 4
      OnExit = cb_depExit
    end
    object edt_type02: TEdit
      Left = 96
      Top = 48
      Width = 121
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 2
    end
    object edt_type01desc: TEdit
      Left = 352
      Top = 16
      Width = 400
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 30
      ParentFont = False
      TabOrder = 1
    end
    object edt_type02desc: TEdit
      Left = 352
      Top = 48
      Width = 400
      Height = 26
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -18
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 40
      ParentFont = False
      TabOrder = 3
    end
  end
  object doc_type: TADOQuery
    Connection = dm.glm
    AfterScroll = doc_typeAfterScroll
    Parameters = <>
    SQL.Strings = (
      'select * from doc_type')
    Left = 184
    Top = 280
    object doc_typeDTYPE01: TWideStringField
      DisplayLabel = #22823#39006#20195#30908
      FieldName = 'DTYPE01'
      Size = 2
    end
    object doc_typeDTYPE02: TWideStringField
      DisplayLabel = #23567#39006#20195#30908
      FieldName = 'DTYPE02'
      Size = 2
    end
    object doc_typeDTYPE_ACTI: TWideStringField
      DisplayLabel = #26159#21542#21487#29992' '
      FieldName = 'DTYPE_ACTI'
      Size = 2
    end
    object doc_typeDTYPE011: TWideStringField
      DisplayLabel = #20027#38542#22823#39006#21517#31281#35498#26126
      DisplayWidth = 15
      FieldName = 'DTYPE011'
      Size = 30
    end
    object doc_typeDTYPE021: TWideStringField
      DisplayLabel = #23376#38542#23567#39006#21517#31281#35498#26126
      DisplayWidth = 20
      FieldName = 'DTYPE021'
      Size = 40
    end
    object doc_typeDTYPE03: TWideStringField
      DisplayLabel = #25215#36774#37096#38272#20195#34399
      FieldName = 'DTYPE03'
      Size = 8
    end
    object doc_typeDTYPE_USER: TWideStringField
      DisplayLabel = #30064#21205#20316#26989#32773
      FieldName = 'DTYPE_USER'
      Size = 8
    end
    object doc_typeDTYPE_DATE: TDateTimeField
      DisplayLabel = #30064#21205#26085#26399
      DisplayWidth = 12
      FieldName = 'DTYPE_DATE'
    end
    object doc_typeDTYPE_FAC: TWideStringField
      FieldName = 'DTYPE_FAC'
      Size = 10
    end
  end
  object DataSource1: TDataSource
    DataSet = doc_type
    Left = 120
    Top = 272
  end
  object PopupMenu1: TPopupMenu
    Left = 312
    Top = 24
  end
end
