Public Sub zwLogIn()

  Dim result As String
  Dim myURL As String, postData As String
  Dim winHttpReq As Object
  Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

  myURL = "https://api.zipwhip.com/user/login"
  postData = "username={userPhoneNumber}&password={userPassword}"

  winHttpReq.Open "POST", myURL, False
  winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  winHttpReq.Send (postData)

  result = winHttpReq.responseText
  Debug.Print (result)

End Sub
