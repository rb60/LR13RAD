object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 473
  ClientWidth = 493
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Menu = MainMenu1
  TextHeight = 15
  object SpeedButton1: TSpeedButton
    Left = 8
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 45
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton2Click
  end
  object SpeedButton3: TSpeedButton
    Left = 74
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton3Click
  end
  object SpeedButton4: TSpeedButton
    Left = 119
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton4Click
  end
  object SpeedButton5: TSpeedButton
    Left = 148
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton5Click
  end
  object SpeedButton6: TSpeedButton
    Left = 191
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton6Click
  end
  object SpeedButton7: TSpeedButton
    Left = 239
    Top = 32
    Width = 23
    Height = 22
    OnClick = SpeedButton7Click
  end
  object OleContainer1: TOleContainer
    Left = 0
    Top = 72
    Width = 493
    Height = 401
    Align = alBottom
    Anchors = [akLeft, akTop, akRight, akBottom]
    Caption = 'OleContainer1'
    TabOrder = 0
  end
  object ComboBox1: TComboBox
    Left = 268
    Top = 34
    Width = 217
    Height = 23
    TabOrder = 1
    Text = 'ComboBox1'
  end
  object OpenDialog1: TOpenDialog
    Left = 24
    Top = 80
  end
  object SaveDialog1: TSaveDialog
    Left = 104
    Top = 80
  end
  object MainMenu1: TMainMenu
    Left = 184
    Top = 80
    object Insert1: TMenuItem
      Caption = 'Insert'
      OnClick = SpeedButton1Click
    end
    object Insert2: TMenuItem
      Caption = 'Create'
      object Word1: TMenuItem
        Caption = 'Word'
        OnClick = SpeedButton2Click
      end
      object FromFile1: TMenuItem
        Caption = 'From File'
        OnClick = SpeedButton3Click
      end
    end
    object Save1: TMenuItem
      Caption = 'Save'
      object asOLEcontainer1: TMenuItem
        Caption = 'as OLE container'
        OnClick = SpeedButton4Click
      end
      object Doc: TMenuItem
        Caption = 'as Doc'
        OnClick = SpeedButton5Click
      end
    end
    object Save2: TMenuItem
      Caption = 'Open'
      OnClick = SpeedButton6Click
    end
    object DoVerb1: TMenuItem
      Caption = 'DoVerb'
      OnClick = SpeedButton7Click
    end
  end
end
