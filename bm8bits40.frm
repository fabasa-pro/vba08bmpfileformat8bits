VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm8bits40 
   Caption         =   "Imagem BMP 8 bits"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm8bits40.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bm8bits40"
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
        
    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"          O tipo de arquivo ("BM").
    HX = HX & "DA050000"    ' BitmapFileSize         DOUBLE WORD    000005DA = 14 + 40 + 1024 + 420 = 1498 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                               Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                               Reservados (0 byte)
    HX = HX & "36040000"    ' BitmapFileOffBits      DOUBLE WORD    00000436 = 14 + 40 + 1024 = 1078 bytes          O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Info Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 40 bytes (padrão).
    
    HX = HX & "28000000"    ' BitmapInfoSize             DOUBLE WORD    00000028 = 40 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "12000000"    ' BitmapInfoWidth            LONG           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "15000000"    ' BitmapInfoHeight           LONG           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapInfoPlanes           WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "0800"        ' BitmapInfoBitCount         WORD               0008 = 8 bpp        Especifica o número de bits por pixel.
    HX = HX & "00000000"    ' BitmapInfoCompression      DOUBLE WORD    00000000 = 0 nenhuma    Especifica o formato de vídeo compactado. (0 nenhuma)
    HX = HX & "A4010000"    ' BitmapInfoSizeImage        DOUBLE WORD    000001A4 = 420 bytes    Especifica o tamanho da imagem.
    HX = HX & "00000000"    ' BitmapInfoXPelsPerMeter    LONG           00000000 = 0 ppm        Especifica a resolução horizontal do dispositivo de destino para o bitmap. (0 ppm)
    HX = HX & "00000000"    ' BitmapInfoYPelsPerMeter    LONG           00000000 = 0 ppm        Especifica a resolução vertical do dispositivo de destino para o bitmap. (0 ppm)
    HX = HX & "00000000"    ' BitmapInfoClrUsed          DOUBLE WORD    00000000 = 0 atributo   Especifica o número de índices de cores na tabela de cores que são realmente usados pelo bitmap. (0 attribute)
    HX = HX & "00000000"    ' BitmapInfoClrImportant     DOUBLE WORD    00000000 = 0 atributo   Especifica o número de índices de cores que são considerados importantes para exibir o bitmap. (0 attribute)
    
    ' Terceira estrutura 'Palette' só será necessária para bitmaps menores que
    ' 24 bits, quando não for possível inserir as cores RGB ou ARGB de cada
    ' pixel diretamente no bitmap e, como nosso bitmap tem 4 bit e utiliza o
    ' cabeçalho Info/ARGB, ela ocupa 256 cores * 4 bytes = 1024 bytes.
    
    HX = HX & "00000000" & "00008000" & "00800000" & "00808000" & "80000000" & "80008000" & "80800000" & "C0C0C000" & "C0DCC000" & "F0CAA600" & "00204000" & "00206000" & "00208000" & "0020A000" & "0020C000" & "0020E000"
    HX = HX & "00400000" & "00402000" & "00404000" & "00406000" & "00408000" & "0040A000" & "0040C000" & "0040E000" & "00600000" & "00602000" & "00604000" & "00606000" & "00608000" & "0060A000" & "0060C000" & "0060E000"
    HX = HX & "00800000" & "00802000" & "00804000" & "00806000" & "00808000" & "0080A000" & "0080C000" & "0080E000" & "00A00000" & "00A02000" & "00A04000" & "00A06000" & "00A08000" & "00A0A000" & "00A0C000" & "00A0E000"
    HX = HX & "00C00000" & "00C02000" & "00C04000" & "00C06000" & "00C08000" & "00C0A000" & "00C0C000" & "00C0E000" & "00E00000" & "00E02000" & "00E04000" & "00E06000" & "00E08000" & "00E0A000" & "00E0C000" & "00E0E000"
    HX = HX & "40000000" & "40002000" & "40004000" & "40006000" & "40008000" & "4000A000" & "4000C000" & "4000E000" & "40200000" & "40202000" & "40204000" & "40206000" & "40208000" & "4020A000" & "4020C000" & "4020E000"
    HX = HX & "40400000" & "40402000" & "40404000" & "40406000" & "40408000" & "4040A000" & "4040C000" & "4040E000" & "40600000" & "40602000" & "40604000" & "40606000" & "40608000" & "4060A000" & "4060C000" & "4060E000"
    HX = HX & "40800000" & "40802000" & "40804000" & "40806000" & "40808000" & "4080A000" & "4080C000" & "4080E000" & "40A00000" & "40A02000" & "40A04000" & "40A06000" & "40A08000" & "40A0A000" & "40A0C000" & "40A0E000"
    HX = HX & "40C00000" & "40C02000" & "40C04000" & "40C06000" & "40C08000" & "40C0A000" & "40C0C000" & "40C0E000" & "40E00000" & "40E02000" & "40E04000" & "40E06000" & "40E08000" & "40E0A000" & "40E0C000" & "40E0E000"
    HX = HX & "80000000" & "80002000" & "80004000" & "80006000" & "80008000" & "8000A000" & "8000C000" & "8000E000" & "80200000" & "80202000" & "80204000" & "80206000" & "80208000" & "8020A000" & "8020C000" & "8020E000"
    HX = HX & "80400000" & "80402000" & "80404000" & "80406000" & "80408000" & "8040A000" & "8040C000" & "8040E000" & "80600000" & "80602000" & "80604000" & "80606000" & "80608000" & "8060A000" & "8060C000" & "8060E000"
    HX = HX & "80800000" & "80802000" & "80804000" & "80806000" & "80808000" & "8080A000" & "8080C000" & "8080E000" & "80A00000" & "80A02000" & "80A04000" & "80A06000" & "80A08000" & "80A0A000" & "80A0C000" & "80A0E000"
    HX = HX & "80C00000" & "80C02000" & "80C04000" & "80C06000" & "80C08000" & "80C0A000" & "80C0C000" & "80C0E000" & "80E00000" & "80E02000" & "80E04000" & "80E06000" & "80E08000" & "80E0A000" & "80E0C000" & "80E0E000"
    HX = HX & "C0000000" & "C0002000" & "C0004000" & "C0006000" & "C0008000" & "C000A000" & "C000C000" & "C000E000" & "C0200000" & "C0202000" & "C0204000" & "C0206000" & "C0208000" & "C020A000" & "C020C000" & "C020E000"
    HX = HX & "C0400000" & "C0402000" & "C0404000" & "C0406000" & "C0408000" & "C040A000" & "C040C000" & "C040E000" & "C0600000" & "C0602000" & "C0604000" & "C0606000" & "C0608000" & "C060A000" & "C060C000" & "C060E000"
    HX = HX & "C0800000" & "C0802000" & "C0804000" & "C0806000" & "C0808000" & "C080A000" & "C080C000" & "C080E000" & "C0A00000" & "C0A02000" & "C0A04000" & "C0A06000" & "C0A08000" & "C0A0A000" & "C0A0C000" & "C0A0E000"
    HX = HX & "C0C00000" & "C0C02000" & "C0C04000" & "C0C06000" & "C0C08000" & "C0C0A000" & "F0FBFF00" & "A4A0A000" & "80808000" & "0000FF00" & "00FF0000" & "00FFFF00" & "FF000000" & "FF00FF00" & "FFFF0000" & "FFFFFF00"
       
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
    
    Open Project.ThisDocument.Path & "\~$bm8bits40.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm8bits40.bmp")
    
End Sub
