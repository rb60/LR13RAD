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
  TextHeight = 15
  object SpeedButton1: TSpeedButton
    Left = 8
    Top = 8
    Width = 23
    Height = 22
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 45
    Top = 8
    Width = 23
    Height = 22
    OnClick = SpeedButton2Click
  end
  object SpeedButton3: TSpeedButton
    Left = 74
    Top = 8
    Width = 23
    Height = 22
    OnClick = SpeedButton3Click
  end
  object SpeedButton4: TSpeedButton
    Left = 119
    Top = 8
    Width = 23
    Height = 22
  end
  object SpeedButton5: TSpeedButton
    Left = 148
    Top = 8
    Width = 23
    Height = 22
  end
  object SpeedButton6: TSpeedButton
    Left = 191
    Top = 8
    Width = 23
    Height = 22
  end
  object SpeedButton7: TSpeedButton
    Left = 239
    Top = 8
    Width = 23
    Height = 22
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
  object Panel1: TPanel
    Left = 0
    Top = -3
    Width = 493
    Height = 69
    TabOrder = 1
    object ComboBox1: TComboBox
      Left = 268
      Top = 10
      Width = 217
      Height = 23
      TabOrder = 0
      Text = 'ComboBox1'
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 24
    Top = 80
  end
  object SaveDialog1: TSaveDialog
    Left = 104
    Top = 80
  end
end
