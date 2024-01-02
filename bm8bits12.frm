VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm8bits12 
   Caption         =   "Imagem BMP 8 bits"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm8bits12.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bm8bits12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

Option Explicit

Private Sub CommandButton1_Click()
    
    ' Declarações gerais:
    
    Dim HX As String    ' Dados (hexadecimal)
    Dim BT As String    ' Bytes
    Dim i As Integer    ' Índices
    
    ' Primeira estrutura 'Bitmap File Header' contém informações sobre o tipo,
    ' tamanho e layout de um bitmap e ocupa 14 bytes (padrão).
        
    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"        O tipo de arquivo ("BM").
    HX = HX & "BE040000"    ' BitmapFileSize         DOUBLE WORD    000004BE = 14 + 12 + 768 + 420 = 1214 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                             Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                             Reservados (0 byte)
    HX = HX & "1A030000"    ' BitmapFileOffBits      DOUBLE WORD    0000031A = 14 + 12 + 768 = 794 bytes            O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Core Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 12 bytes (padrão).
    
    HX = HX & "0C000000"    ' BitmapCoreSize         DOUBLE WORD    0000000C = 12 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "1200"        ' BitmapCoreWidth        WORD           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "1500"        ' BitmapCoreHeight       WORD           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapCorePlanes       WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "0800"        ' BitmapCoreBitCoun      WORD               0008 = 8 bpp        Especifica o número de bits por pixel.
    
    ' Terceira estrutura 'Palette' só será necessária para bitmaps menores que
    ' 24 bits, quando não for possível inserir as cores RGB ou ARGB de cada
    ' pixel diretamente no bitmap e, como nosso bitmap tem 4 bit e utiliza o
    ' cabeçalho Core/RGB, ela ocupa 256 cores * 3 bytes = 768 bytes.
    
    HX = HX & "000000" & "000080" & "008000" & "008080" & "800000" & "800080" & "808000" & "C0C0C0" & "C0DCC0" & "F0CAA6" & "002040" & "002060" & "002080" & "0020A0" & "0020C0" & "0020E0"
    HX = HX & "004000" & "004020" & "004040" & "004060" & "004080" & "0040A0" & "0040C0" & "0040E0" & "006000" & "006020" & "006040" & "006060" & "006080" & "0060A0" & "0060C0" & "0060E0"
    HX = HX & "008000" & "008020" & "008040" & "008060" & "008080" & "0080A0" & "0080C0" & "0080E0" & "00A000" & "00A020" & "00A040" & "00A060" & "00A080" & "00A0A0" & "00A0C0" & "00A0E0"
    HX = HX & "00C000" & "00C020" & "00C040" & "00C060" & "00C080" & "00C0A0" & "00C0C0" & "00C0E0" & "00E000" & "00E020" & "00E040" & "00E060" & "00E080" & "00E0A0" & "00E0C0" & "00E0E0"
    HX = HX & "400000" & "400020" & "400040" & "400060" & "400080" & "4000A0" & "4000C0" & "4000E0" & "402000" & "402020" & "402040" & "402060" & "402080" & "4020A0" & "4020C0" & "4020E0"
    HX = HX & "404000" & "404020" & "404040" & "404060" & "404080" & "4040A0" & "4040C0" & "4040E0" & "406000" & "406020" & "406040" & "406060" & "406080" & "4060A0" & "4060C0" & "4060E0"
    HX = HX & "408000" & "408020" & "408040" & "408060" & "408080" & "4080A0" & "4080C0" & "4080E0" & "40A000" & "40A020" & "40A040" & "40A060" & "40A080" & "40A0A0" & "40A0C0" & "40A0E0"
    HX = HX & "40C000" & "40C020" & "40C040" & "40C060" & "40C080" & "40C0A0" & "40C0C0" & "40C0E0" & "40E000" & "40E020" & "40E040" & "40E060" & "40E080" & "40E0A0" & "40E0C0" & "40E0E0"
    HX = HX & "800000" & "800020" & "800040" & "800060" & "800080" & "8000A0" & "8000C0" & "8000E0" & "802000" & "802020" & "802040" & "802060" & "802080" & "8020A0" & "8020C0" & "8020E0"
    HX = HX & "804000" & "804020" & "804040" & "804060" & "804080" & "8040A0" & "8040C0" & "8040E0" & "806000" & "806020" & "806040" & "806060" & "806080" & "8060A0" & "8060C0" & "8060E0"
    HX = HX & "808000" & "808020" & "808040" & "808060" & "808080" & "8080A0" & "8080C0" & "8080E0" & "80A000" & "80A020" & "80A040" & "80A060" & "80A080" & "80A0A0" & "80A0C0" & "80A0E0"
    HX = HX & "80C000" & "80C020" & "80C040" & "80C060" & "80C080" & "80C0A0" & "80C0C0" & "80C0E0" & "80E000" & "80E020" & "80E040" & "80E060" & "80E080" & "80E0A0" & "80E0C0" & "80E0E0"
    HX = HX & "C00000" & "C00020" & "C00040" & "C00060" & "C00080" & "C000A0" & "C000C0" & "C000E0" & "C02000" & "C02020" & "C02040" & "C02060" & "C02080" & "C020A0" & "C020C0" & "C020E0"
    HX = HX & "C04000" & "C04020" & "C04040" & "C04060" & "C04080" & "C040A0" & "C040C0" & "C040E0" & "C06000" & "C06020" & "C06040" & "C06060" & "C06080" & "C060A0" & "C060C0" & "C060E0"
    HX = HX & "C08000" & "C08020" & "C08040" & "C08060" & "C08080" & "C080A0" & "C080C0" & "C080E0" & "C0A000" & "C0A020" & "C0A040" & "C0A060" & "C0A080" & "C0A0A0" & "C0A0C0" & "C0A0E0"
    HX = HX & "C0C000" & "C0C020" & "C0C040" & "C0C060" & "C0C080" & "C0C0A0" & "F0FBFF" & "A4A0A0" & "808080" & "0000FF" & "00FF00" & "00FFFF" & "FF0000" & "FF00FF" & "FFFF00" & "FFFFFF"
       
    ' Quarta estrutura 'Bitmap' contém todos os pixels extrudados em uma matriz
    ' de coluna e linha, onde temos linhas de 0 a 20 = 21 de altura e 18 na
    ' largura, em partes de 32 bits, por esse motivo completamos com 0 (zero)
    ' até obter os completos 32 bits, ela ocupa 21 linhas * 20 bytes = 420 bytes.
        
    '       32 bits     32 bits     32 bits     32 bits     32 bits
    '     ----------- ----------- ----------- ----------- -----------
    '  0: FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00
    '  1: FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00
    '  2: FF FF FF FF FF FF FF 00 00 00 00 FF FF FF FF FF FF FF 00 00
    '  3: FF FF FF FF FF 00 00 FB FB FB FB 00 00 FF FF FF FF FF 00 00
    '  4: FF FF FF FF 00 FB FB FB FB FB FB FB FB 00 FF FF FF FF 00 00
    '  5: FF FF FF FF 00 FB FB FB FB FB FB FB FB 00 FF FF FF FF 00 00
    '  6: FF FF FF 00 FB FF FB FB FB FB FB FB FB FB 00 FF FF FF 00 00
    '  7: FF FF FF 00 FF FF FF FF FB FB FB FB FB FF 00 FF FF FF 00 00
    '  8: FF FF FF 00 FF FF FF FF FF FF FF FF FF FF 00 FF FF FF 00 00
    '  9: FF FF FF FF 00 FF FF 00 FF FF 00 FF FF 00 FF FF FF FF 00 00
    ' 10: FF FF FF FF FF 00 FF 00 FF FF 00 FF 00 FF FF FF FF FF 00 00
    ' 11: FF FF FF FF 00 F9 00 FF FF FF FF 00 F9 00 FF FF FF FF 00 00
    ' 12: FF FF FF 00 FF F9 F9 00 00 00 00 F9 F9 FF 00 FF FF FF 00 00
    ' 13: FF FF 00 FF FF 00 F9 F9 F9 F9 F9 F9 00 FF FF 00 FF FF 00 00
    ' 14: FF FF 00 FF FF 00 F9 F9 F9 F9 F9 F9 00 FF FF 00 FF FF 00 00
    ' 15: FF FF FF 00 00 FE 00 00 00 00 00 00 FE 00 00 FF FF FF 00 00
    ' 16: FF FF FF FF 00 FE FE FE FE FE FE FE FE 00 FF FF FF FF 00 00
    ' 17: FF FF FF FF 00 FC FC FC 00 00 FC FC FC 00 FF FF FF FF 00 00
    ' 18: FF FF FF FF FF 00 00 00 FF FF 00 00 00 FF FF FF FF FF 00 00
    ' 19: FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00
    ' 20: FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00
    
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 20:                                                       00 00
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 19:                                                       00 00
    HX = HX & "FFFFFFFFFF000000FFFF000000FFFFFFFFFF0000"    ' 18:                00 00 00       00 00 00                00 00
    HX = HX & "FFFFFFFF00FCFCFC0000FCFCFC00FFFFFFFF0000"    ' 17:             00 FC FC FC 00 00 FC FC FC 00             00 00
    HX = HX & "FFFFFFFF00FEFEFEFEFEFEFEFE00FFFFFFFF0000"    ' 16:             00 FE FE FE FE FE FE FE FE 00             00 00
    HX = HX & "FFFFFF0000FE000000000000FE0000FFFFFF0000"    ' 15:          00 00 FE 00 00 00 00 00 00 FE 00 00          00 00
    HX = HX & "FFFF00FFFF00F9F9F9F9F9F900FFFF00FFFF0000"    ' 14:       00       00 F9 F9 F9 F9 F9 F9 00       00       00 00
    HX = HX & "FFFF00FFFF00F9F9F9F9F9F900FFFF00FFFF0000"    ' 13:       00       00 F9 F9 F9 F9 F9 F9 00       00       00 00
    HX = HX & "FFFFFF00FFF9F900000000F9F9FF00FFFFFF0000"    ' 12:          00    F9 F9 00 00 00 00 F9 F9    00          00 00
    HX = HX & "FFFFFFFF00F900FFFFFFFF00F900FFFFFFFF0000"    ' 11:             00 F9 00             00 F9 00             00 00
    HX = HX & "FFFFFFFFFF00FF00FFFF00FF00FFFFFFFFFF0000"    ' 10:                00    00       00    00                00 00
    HX = HX & "FFFFFFFF00FFFF00FFFF00FFFF00FFFFFFFF0000"    '  9:             00       00       00       00             00 00
    HX = HX & "FFFFFF00FFFFFFFFFFFFFFFFFFFF00FFFFFF0000"    '  8:          00                               00          00 00
    HX = HX & "FFFFFF00FFFFFFFFFBFBFBFBFBFF00FFFFFF0000"    '  7:          00             FB FB FB FB FB    00          00 00
    HX = HX & "FFFFFF00FBFFFBFBFBFBFBFBFBFB00FFFFFF0000"    '  6:          00 FB    FB FB FB FB FB FB FB FB 00          00 00
    HX = HX & "FFFFFFFF00FBFBFBFBFBFBFBFB00FFFFFFFF0000"    '  5:             00 FB FB FB FB FB FB FB FB 00             00 00
    HX = HX & "FFFFFFFF00FBFBFBFBFBFBFBFB00FFFFFFFF0000"    '  4:             00 FB FB FB FB FB FB FB FB 00             00 00
    HX = HX & "FFFFFFFFFF0000FBFBFBFB0000FFFFFFFFFF0000"    '  3:                00 00 FB FB FB FB 00 00                00 00
    HX = HX & "FFFFFFFFFFFFFF00000000FFFFFFFFFFFFFF0000"    '  2:                      00 00 00 00                      00 00
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  1:                                                       00 00
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  0:                                                       00 00
    
    ' Salvar arquivo bitmap 256 cores (*.bmp;*.dib).
    
    Open Project.ThisDocument.Path & "\~$bm8bits12.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm8bits12.bmp")
    
End Sub
