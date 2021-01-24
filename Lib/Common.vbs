Option Explicit

		' ------------------------------------------------------------
		' ���
		' ------------------------------------------------------------
		Public Sub Substitution(ByRef distination, ByRef source)
			If IsObject(source) Then
				Set distination = source
			Else
				distination = source
			End If
		End Sub

		Public Function IIF(ByVal expr, ByVal TrueResult, ByVal FalseResult)
			If expr Then
				Substitution IIF, TrueResult
			Else
				Substitution IIF, FalseResult
			End If
		End Function

		Public Function IsString(ByVal var)
			If Not (IsObject(var) Or IsNull(var)) Then
				If IsNumeric(var) Then
					IsString = ((var+1)-1) <> var
				Else
					IsString = True
				End If
			End If
		End Function

		' ------------------------------------------------------------
		' YYYYMMDD�ҏW
		' ------------------------------------------------------------
		Function YYYYMMDD(dt)
			Dim y, m, d
			y = "0" & Year(dt)
			m = "0" & Month(dt)
			d = "0" & Day(dt)
			YYYYMMDD = Right(y,4) & Right(m,2) & Right(d,2)
		End Function

		' ------------------------------------------------------------
		' HHNNSS�ҏW
		' ------------------------------------------------------------
		Function HHNNSS(dt)
			Dim h, n, s
			h = "0" & hour(dt)
			n = "0" & Minute(dt)
			s = "0" & Second(dt)
			HHNNSS = Right(h,2) & Right(n,2) & Right(s,2)
		End Function

		' ------------------------------------------------------------
		' YYYYMMDDHHMMSS
		' ------------------------------------------------------------
		Function YMDHMS(dt)
			YMDHMS = YYYYMMDD(dt) & HHNNSS(dt)
		End Function

		' ------------------------------------------------------------
		' "YYYY/MM/DD HH:MM:SS"
		' ------------------------------------------------------------
		Function ToDTS(dt)
			Dim src, d, t
			src = YYYYMMDD(dt)
			d = Mid(src, 1, 4) & "/" & Mid(src, 5, 2) & "/" & Mid(src, 7, 2)
			src = HHNNSS(dt)
			t = Mid(src, 9, 2) & ":" & Mid(src, 11, 2) & ":" & Mid(src, 13, 2)
			ToDTS = d & " " & t
		End Function

		' ----------------------------------------
		' 10���̐���������
		' ----------------------------------------
		Function Numeric10(num)
			Numeric10 = Right("0000000000" & CStr(num), 10)
		End Function

		' ------------------------------------------------------------
		' �����ҏW
		' ------------------------------------------------------------
		Function MessageFormat(format, arr)
			Dim dic, par
			Set dic = CreateObject("Scripting.Dictionary")
			Set par = CreateObject("Scripting.Dictionary")

			' �p�����[�^
			For idx=LBound(arr) To UBound(arr)
				par.Add Cstr(idx+1), arr(idx)
			Next

			' ���K�\���ŏ����u���b�N�����
			Dim regEx, Matches, text
			Set regEx = New RegExp
			regEx.Pattern = "\$\{([0-9]+)[:]?([0-9]*)\}"
			regEx.IgnoreCase = True
			regEx.Global = True
			text = sun(format)
			Set Matches = regEx.Execute(text)

			' �����u���b�N���p�����[�^�Œu��
			Dim RetStr, pos, Match, SubMatch, fmt, idx, key
			RetStr = ""
			pos = 1
			For Each Match in Matches
				' �e�L�X�g�u���b�N��ǉ�
				RetStr = RetStr & unsun(Mid(text, pos, Match.FirstIndex+1 - pos))
				pos = Match.FirstIndex + Match.Length + 1
				' �����ɏ]���ăp�����[�^���t�H�[�}�b�g
				key = Match.Value
				If Not dic.Exists(key) Then
					Set fmt = new formatter
					fmt.Init Match, par
					dic.Add key, fmt
				Else
					Set fmt = dic.Item(key)
				End IF
				' �t�H�[�}�b�g�ς݃p�����[�^��ǉ�
				RetStr = RetStr & fmt.value
			Next
			' �c�����e�L�X�g�u���b�N��ǉ�
			RetStr = RetStr & unsun(Mid(text, pos))
			MessageFormat = RetStr
		End Function

' --------------------------------------------------------------------------------

		'Const fmt = "<${1} ${2:2} ${3}\n\t${1}\${3}\\${2:4} ${4>"
		'MsgBox(fmt & vbLf & MessageFormat(fmt, Array("aaa","bbb","ccc")))
		Private Function sun(str)
			sun = str
			sun = Replace(sun, "&", "&amp;")
			sun = Replace(sun, "<", "&lt;")
			sun = Replace(sun, ">", "&gt;")
			sun = Replace(sun, "\\", "<ESC>")
			sun = Replace(sun, "\n", "<RET>")
			sun = Replace(sun, "\t", "<TAB>")
			sun = Replace(sun, "\$", "<DOL>")
		End Function

		Private Function unsun(str)
			unsun = str
			unsun = Replace(unsun, "<DOL>", "$")
			unsun = Replace(unsun, "<TAB>", vbTab)
			unsun = Replace(unsun, "<RET>", vbCrLf)
			unsun = Replace(unsun, "<ESC>", "\")
			unsun = Replace(unsun, "&gt;", ">")
			unsun = Replace(unsun, "&lt;", "<")
			unsun = Replace(unsun, "&amp;", "&")
		End Function

' --------------------------------------------------------------------------------

	Class formatter
		Dim key
		Dim paramnum
		Dim width
		Dim value

		Sub Init(Match, params)
			key = Match.Value
			paramnum = CInt("0" & Match.SubMatches(0))
			width = CInt("0" & Match.SubMatches(1))
			Dim prmkey
			prmkey = CStr(paramnum)
			If params.Exists(prmkey) Then
				' �Ή��o�����[�^����
				value = params.Item(prmkey)
				if width>LenB(value) then
					value = Space(width-LenB(value)) & value
				End If
			Else
				' �Ή��o�����[�^�Ȃ�
				value = Space(width)
			End IF
		End Sub
	End Class

' --------------------------------------------------------------------------------

		' ------------------------------------------------------------
		' �t�@�C���� UTF-8 �G���R�[�f�B���O�ŕۑ����܂�
		' (���l�^) http://d.hatena.ne.jp/replication/20091117/1258418243
		' ------------------------------------------------------------
		Sub SaveAsUTF8(filename, text)
			On Error Resume Next

			' ADODB.Stream�̃��[�h
			Dim adTypeBinary : adTypeBinary = 1
			Dim adTypeText : adTypeText = 2
			Dim adSaveCreateOverWrite : adSaveCreateOverWrite = 2

			' ADODB.Stream���쐬
			Dim pre : Set pre = CreateObject("ADODB.Stream")
			' �ŏ��̓e�L�X�g���[�h��UTF-8�ŏ�������
			pre.Type = adTypeText
			pre.Charset = "UTF-8"
			pre.Open()
			pre.WriteText(text)
			' �o�C�i�����[�h�ɂ��邽�߂�Position����x0�ɖ߂�
			' Read���邽�߂ɂ̓o�C�i���^�C�v�łȂ��Ƃ����Ȃ�
			pre.Position = 0
			pre.Type = adTypeBinary
			' Position��3�ɂ��Ă���ǂݍ��ނ��Ƃōŏ���3�o�C�g���X�L�b�v����
			' �܂�BOM���X�L�b�v���܂�
			pre.Position = 3
			Dim bin : bin = pre.Read()
			pre.Close()

			' �ǂݍ��񂾃o�C�i���f�[�^���o�C�i���f�[�^�Ƃ��ăt�@�C���ɏo�͂���
			' �����͈�ʓI�ȏ������Ȃ̂Ő������ȗ�
			Dim stm : Set stm = CreateObject("ADODB.Stream")
			stm.Type = adTypeBinary
			stm.Open()
			stm.Write(bin)
			stm.SaveToFile filename, adSaveCreateOverWrite ' force overwrite
			stm.Close()

		End Sub
