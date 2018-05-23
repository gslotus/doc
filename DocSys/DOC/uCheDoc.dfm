object fmCheDep: TfmCheDep
  Left = 128
  Top = 132
  Width = 928
  Height = 480
  Caption = #27298#26597#37096#38272#30064#21205
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 912
    Height = 65
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 24
      Width = 96
      Height = 16
      Caption = #27298#26597#26085#26399#21312#38291
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 280
      Top = 23
      Width = 8
      Height = 16
      Caption = '~'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object dtp_date1: TDateTimePicker
      Left = 128
      Top = 23
      Width = 145
      Height = 24
      CalAlignment = dtaLeft
      Date = 40264.4035225579
      Time = 40264.4035225579
      DateFormat = dfShort
      DateMode = dmComboBox
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      Kind = dtkDate
      ParseInput = False
      ParentFont = False
      TabOrder = 0
    end
    object dtp_date2: TDateTimePicker
      Left = 304
      Top = 23
      Width = 161
      Height = 24
      CalAlignment = dtaLeft
      Date = 40264.4035966898
      Time = 40264.4035966898
      DateFormat = dfShort
      DateMode = dmComboBox
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      Kind = dtkDate
      ParseInput = False
      ParentFont = False
      TabOrder = 1
    end
    object Button1: TButton
      Left = 515
      Top = 21
      Width = 85
      Height = 33
      Caption = #38283#22987#27298#26597
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
      OnClick = Button1Click
    end
  end
  object dep_new: TDBGrid
    Left = 0
    Top = 65
    Width = 912
    Height = 375
    Align = alClient
    DataSource = ds_new
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
        FieldName = 'CPF02'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CQA03'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CQA04'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'BDEP'
        Title.Alignment = taCenter
        Width = 180
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CQA05'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'ADEP'
        Title.Alignment = taCenter
        Width = 195
        Visible = True
      end>
  end
  object ds_new: TDataSource
    DataSet = ado_new
    Left = 120
    Top = 248
  end
  object ado_new: TADOQuery
    Connection = dm.glm
    Parameters = <>
    SQL.Strings = (
      
        'select cpf02,CQA03,CQA04,CQA05 from cqa_file,cpf_file,cqaa_file ' +
        'where cqa01 = cqaa01 and cpf01 = cqa03'
      ' and cqa06 ='#39'5'#39)
    Left = 288
    Top = 240
    object ado_newCPF02: TWideStringField
      DisplayLabel = #21729#24037#22995#21517
      FieldName = 'CPF02'
      Size = 10
    end
    object ado_newCQA03: TWideStringField
      DisplayLabel = #24037#34399
      FieldName = 'CQA03'
      Size = 8
    end
    object ado_newCQA04: TWideStringField
      DisplayLabel = #30064#21205#21069#37096#38272
      FieldName = 'CQA04'
      Size = 12
    end
    object ado_newBDEP: TStringField
      DisplayLabel = #37096#38272#21517#31281
      FieldName = 'BDEP'
    end
    object ado_newCQA05: TWideStringField
      DisplayLabel = #30064#21205#24460#37096#38272
      FieldName = 'CQA05'
      Size = 12
    end
    object ado_newADEP: TStringField
      DisplayLabel = #37096#38272#21517#31281
      FieldName = 'ADEP'
    end
  end
end
