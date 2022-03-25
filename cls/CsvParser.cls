Option Explicit

' �N�H�[�g����
Private m_quote As String
' �w�b�_�s�������̃t���O
Private m_hasHeader As Boolean
' �t�@�C����
Private m_fileNum As Integer
' �s�̖������̃t���O
Private m_endLine As Boolean
' �t�@�C���̖����܂œǂݏI��������̃t���O
Private m_endOfFile As Boolean
' �t�@�C�����J���Ă��邩�̃t���O
Private m_fileIsOpen As Boolean

' *************************************************
' * �R���X�g���N�^
' *
' *************************************************
Private Sub Class_Initialize()
    m_fileIsOpen = False
End Sub

' *************************************************
' * �f�X�g���N�^
' *
' *************************************************
Private Sub Class_Terminate()
    CloseFile
End Sub

' *************************************************
' * �t�@�C�����e��Ǎ���
' *
' * Params
' * ------
' * filename: String
' *     �Ǎ��ݑΏۃt�@�C���̃p�X
' * hasHeader: Boolean
' *     1�s�ڂ��w�b�_�Ƃ��Ĉ������̃t���O
' * quote: String
' *     �͂ݕ���
' *
' * Return
' * ------
' * �t�@�C�����e: Collection
' *     �e�v�f�̓w�b�_�������͘A�Ԃ��L�[�ɂ��������I�u�W�F�N�g
' *
' *************************************************
Public Function readFile(filename As String, Optional hasHeader As Boolean = False, Optional quote As String = """") As Collection
    ' �����o�ݒ�
    m_quote = quote
    m_hasHeader = hasHeader

    ' �t�@�C���I�[�v��
    Call openFile(filename)

    ' ���e��Ǎ�
    Set readFile = readAllValues()
End Function

' *************************************************
' * �t�@�C�����I�[�v������
' *
' * Params
' * ------
' * filename: String
' *     �Ǎ��ݑΏۃt�@�C���̃p�X
' *
' *************************************************
Private Function openFile(filename As String)
    m_fileNum = FreeFile
    Open filename For Input Access Read As m_fileNum
    m_fileIsOpen = True
End Function

' *************************************************
' * �t�@�C�����N���[�Y����
' *
' * Params
' * ------
' * filename: String
' *     �Ǎ��ݑΏۃt�@�C���̃p�X
' *
' *************************************************
Function CloseFile()
    If m_fileIsOpen Then
        Close m_fileNum
        m_fileIsOpen = False
    End If
End Function

' *************************************************
' * �t�@�C���̑S�Ă̓��e��ǂݍ���
' *
' * Params
' * ------
' * filename: String
' *     �Ǎ��ݑΏۃt�@�C���̃p�X
' *
' * Return
' * ------
' * �t�@�C�����e: Collection
' *
' *************************************************
Private Function readAllValues() As Collection
    Set readAllValues = New Collection
    Dim values As Object
    Dim lineValues As Collection
    Dim columns As Collection
    Dim idx As Integer
    
    If m_hasHeader Then
        Set columns = readLine()
    End If

    Do Until EOF(m_fileNum)
        ' 1�s���Ǎ��݁A�����ɂ���
        Set values = CreateObject("Scripting.Dictionary")
        Set lineValues = readLine()

        If m_hasHeader Then
            ' �w�b�_������ꍇ�A�w�b�_�̊e���ږ����L�[�Ƃ��������Ƃ���
            For idx = 1 To columns.Count
                Call values.Add(columns.item(idx), lineValues.item(idx))
            Next
        Else
            ' �w�b�_���Ȃ��ꍇ�A�A�Ԃ��L�[�Ƃ���
            For idx = 1 To lineValues.Count
                Call values.Add(idx, lineValues.item(idx))
            Next
        End If
        Call readAllValues.Add(values)
    Loop
End Function

' *************************************************
' * �t�@�C������1�s�Ǎ���
' *
' * Return
' * ------
' * �s���e: Collection
' *
' *************************************************
Private Function readLine() As Collection
    Set readLine = New Collection
    m_endLine = False

    Do Until m_endLine
        readLine.Add (readValue())
    Loop
End Function

' *************************************************
' * �t�@�C������1���ړǍ���
' *
' * Return
' * ------
' * ���ړ��e: String
' *
' *************************************************
Private Function readValue() As String
    Dim buffer As String
    buffer = ""

    Do Until isCloseValue(buffer)
        buffer = buffer & Input(1, m_fileNum)
    Loop

    readValue = extractValue(buffer)
End Function

' *************************************************
' * ���ڂ̏I���܂œǍ��܂�Ă��邩���肷��
' *
' * Return
' * ------
' * �Ǎ��܂�Ă��邩�̃t���O: Boolean
' *
' *************************************************
Private Function isCloseValue(buffer) As Boolean
    Select Case Right(buffer, 1)
        Case ",", vbLf
            isCloseValue = True
        Case Else
            isCloseValue = False
    End Select
End Function

' *************************************************
' * �����񂩂獀�ڒl�����݂̂𒊏o����
' *
' * Return
' * ------
' * ���o�����l: String
' *
' *************************************************
Private Function extractValue(target As String) As String
    If Left(target, 1) = m_quote Then
        ' �N�H�[�g�͂ݗL��
        If Right(target, 2) = m_quote & "," Then
            extractValue = Mid(target, 2, Len(target) - 3)
        ElseIf Right(target, 2) = m_quote & vbLf Then
            extractValue = Mid(target, 2, Len(target) - 3)
            m_endLine = True
        ElseIf Right(target, 3) = m_quote & vbCr & vbLf Then
            extractValue = Mid(target, 2, Len(target) - 4)
            m_endLine = True
        Else
            Err.Raise 9999, Description:="�t�@�C���t�H�[�}�b�g�s��"
        End If
    Else
        ' �N�H�[�g�͂ݖ���
        If Right(target, 1) = "," Then
            extractValue = Mid(target, 1, Len(target) - 1)
        ElseIf Right(target, 1) = vbLf Then
            extractValue = Mid(target, 1, Len(target) - 1)
            m_endLine = True
        ElseIf Right(target, 2) = vbCr & vbLf Then
            extractValue = Mid(target, 1, Len(target) - 2)
            m_endLine = True
        Else
            Err.Raise 9999, Description:="�t�@�C���t�H�[�}�b�g�s��"
        End If
    End If
End Function