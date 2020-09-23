<div align="center">

## Read Webpages with Inet \(shows all of the source\)


</div>

### Description

This code uses the MSInet.ocx control to read webpages. Since Inet has a bug with it's OpenURL function, it uses the GetChunk method instead.
 
### More Info
 
Requirements: Richtextbox.ocx, MSInet.ocx

The code returns all of the webpage's source.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Gates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-gates.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-gates-read-webpages-with-inet-shows-all-of-the-source__1-8967/archive/master.zip)





### Source Code

```
Screen.MousePointer = 11
Dim ReturnStr As String
RichTextBox1.Text = Inet1.OpenURL("http://www.planet-source-code.com", icString)
ReturnStr = Inet1.GetChunk(2048, icString)
Do While Len(ReturnStr) <> 0
 DoEvents
 RichTextBox1.Text = RichTextBox1.Text & ReturnStr
 ReturnStr = Inet1.GetChunk(2048, icString)
Loop
Screen.MousePointer = 0
```

