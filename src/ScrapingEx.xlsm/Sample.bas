Attribute VB_Name = "Sample"
''' --------------------------------------------------------
'''  FILE    : Sample.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

''' ���z�[�Ō�������
Public Sub sample_yahoo()

    ''' �X�N���C�s���O�d�����g����悤�ɂ��܂�
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' ���z�[�̃z�[���y�[�W���J���܂�
    doc.GotoPage "https://www.yahoo.co.jp/"

    ''' �������� VBA �Ɠ��͂��܂�
    doc.ID("srchtxt").FieldValue "VBA"

    ''' �����{�^���������܂�
    doc.ID("srchbtn").Click

End Sub

''' �O�[�O���Ō�������
Public Sub sample_google()

    ''' �X�N���C�s���O�d�����g����悤�ɂ��܂�
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' �O�[�O���̃z�[���y�[�W���J���܂�
    doc.GotoPage "https://www.google.com/"

    ''' �������� VBA �Ɠ��͂��܂�
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.RNNXgb > div > div.a4bIc > input").FieldValue "VBA"

    ''' �����{�^���������܂�
    doc.At_CSS("#tsf > div:nth-child(2) > div > div.FPdoLc.VlcLAe > center > input.gNO89b").Click

End Sub

''' ���g�U�̍ŐV�̌��ʂ��擾����
Public Sub Sample_Loto6()

    ''' �X�N���C�s���O�d�����g����悤�ɂ��܂�
    Dim doc As ScrapingEx: Set doc = New ScrapingEx

    ''' ���g�U�̃z�[���y�[�W���J���܂�
    doc.GotoPage "https://www.mizuhobank.co.jp/retail/takarakuji/loto/loto6/index.html"

    ''' ���g�U�̍ŐV�̌��ʕ\�̃L�����[�I�[�o�[�̋��z�̃Z�����󔒂łȂ���ԂɂȂ�܂ő҂��܂�
    Dim selector As String: selector = "#mainCol > article > section > section > section > div > div.sp-none > table:nth-child(1) > tbody > tr:nth-child(10) > td > strong"
    doc.Until_TextMatches selector, "[^ \t\n\r\f]"

    ''' ���g�U�̍ŐV�̌��ʕ\��z��ɂ��܂�
    Dim tableArr As Variant
    tableArr = ArrTable(doc.CSS("table.typeTK").Index(0).RowTable, True)(1)

    ''' �C�~�f�B�G�C�g�E�B���h�E�ɂĎ擾�����f�[�^��\�����܂�
    Dim v
    For Each v In tableArr
        Debug.Print Join(v, " ")
    Next v

    ''' �u���E�U�i�h�d�j��Еt���܂�
    doc.Quit
    Set doc = Nothing

End Sub
