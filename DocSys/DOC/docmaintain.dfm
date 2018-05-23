object fmMaintain: TfmMaintain
  Left = 193
  Top = 107
  Width = 696
  Height = 480
  BorderIcons = [biSystemMenu]
  Caption = #20195#30908#32173#35703
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 106
  TextHeight = 13
  object Label2: TLabel
    Left = 8
    Top = 40
    Width = 172
    Height = 16
    Caption = 'YY'#24180#22577' MM'#26376#22577'  SS'#23395#22577
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 56
    Width = 661
    Height = 389
    Align = alBottom
    DataSource = DataSource1
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -16
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
        FieldName = 'DTYPE03'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DTYPE04'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DTYPE06'
        Visible = True
      end>
  end
  object DBNavigator1: TDBNavigator
    Left = 0
    Top = 0
    Width = 661
    Height = 25
    DataSource = DataSource1
    Align = alTop
    TabOrder = 1
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 200
    Top = 80
  end
  object ADOQuery1: TADOQuery
    Connection = dm.gshank
    Parameters = <>
    SQL.Strings = (
      'select * from doc_type')
    Left = 368
    Top = 184
    object ADOQuery1DTYPE01: TWideStringField
      DisplayLabel = #20195#30908#22823#39006
      FieldName = 'DTYPE01'
      Size = 2
    end
    object ADOQuery1DTYPE02: TWideStringField
      DisplayLabel = #20195#30908#23567#39006
      FieldName = 'DTYPE02'
      Size = 2
    end
    object ADOQuery1DTYPE03: TWideStringField
      DisplayLabel = #20195#30908#35498#26126
      DisplayWidth = 30
      FieldName = 'DTYPE03'
      Size = 100
    end
    object ADOQuery1DTYPE04: TWideStringField
      DisplayLabel = #39006#22411
      FieldName = 'DTYPE04'
      Size = 2
    end
    object ADOQuery1DTYPE05: TWideStringField
      DisplayLabel = #26159#21542#26367#25563
      FieldName = 'DTYPE05'
      Size = 2
    end
    object ADOQuery1DTYPE06: TWideStringField
      DisplayLabel = #26159#21542#21487#29992
      FieldName = 'DTYPE06'
      Size = 2
    end
  end
end
