Attribute VB_Name = "bsEngineSign"
Option Explicit
Public ViriIconID() As String, ViriiconNa() As String
Public Const IntViriIconID As String = "1F1C9B9|20938B2|19F4ED6|133BE0B|18EDEAE|1EF89C2|24563C4|1B2DB74|208EA72|22A064D|19B64EE|1D4B7E1|2087762|29C7258|1B18705|1B5FCAB|126D4CF|1C58E5C|15D7730|1FB82B7|112763E|2165AF9|25F46BE|206556B|22A8D69|19237F8|15022B4|1D8B4EB|1DBC1EA|2333F5D|1F37C2F|1C9CCA4|1DFDFB4|1C1283E|1F6598C|27F4C1A|22F92E0|191DBDC|27BFE4A|20E0907|27C16B8|1EE74B8|242721D|22CF2FB|22D52E3|1E983D3|22D7F29|22D3504|22D4091|1EB03A3|22DB4CC|1EA2FEF|1EA77DE|1E99419|22C27CC|B64AAA|1E17D43|20938B2|20938B2|1EF89C2|1EF02E3|22D07EF|22F92E0|7D59C2|18329B6|21D1F10|2577E5B|1DE28E2"
Public Const IntViriVariantName As String = "Aduhai|Worm.Folder|Kangen|Apel|Apel|Brontok|Ascribes|Codex|Brontok|Cyrax|Cyrax|Decoil|Rolog|Ego|FluBurung|Gelas|Imelda|Imelda|Iwing|Jablay|KamaSutra|Worm.Documen|Mazda|MySong|Nahital|Netsky|Riyani|Nimda|Nukedevil|Parayrontok|Peta|Pluto|Pluto|Polyface|Provisioning|Renova|Worm.Folder|Tinutuan|Tsunami|Wukill|Junx.oB|Amburadul|Worm.Folder2|W32.AAC|W32.APE|W32.asF|W32.avi|W32.flac|W32.m4a|W32.mpc|W32.mpg|W32.ogg|W32.ogm|W32.vid|W32.wmv|W32.wv|W32.ffd|W32.dir5|W32.dir6|W32.dir7|IE.dir8|W32.wnam|W32.txt|W32.ac|W32.as|W32.docXP|W32.dir10|W32.pic"

Public jmlVER As Long
Public VirVERid() As String, virVERname() As String
Public Const VirVER As String = "VIDEO|KB"
Public Const VirVERN As String = "Fake.Video|Fake.File"

Public jmlVPE As Long
Public VirPEid() As String, virPEname() As String
Public Const virPE As String = "59BFFA310F12FC|440FAC91818E1000|477A363610F14570|477A359210FCA80|477A355310FA8F0|477A350F10FA9B0|477A347F10F98A0|477A341910FBCB0|4CD1C353210E1184|4CD1C353210E1184|4910D53B10F113C|4882DDA210F1100|48BE596510F1100|464046BF10F1D64|4640403610F1C80|4A1AB0DE10F21FC0|2A425E19818E59654|4871B32A10F1394|486B70FF210218D5C|38730CAE10E1128|4AC11C2710F1100|4C2B232A10F1710|4B706D391021B6F|36EEE8D6210E434B|3FAEBF54210E486B|488D7F37210ED41B|4C457CF610F11D4|34497E1710F11DC|3CA2BE0010F112C|4CD1C353210E1184"
Public Const virPEN As String = "Troj.Dropper.GEN|FakeAV.XPack|Lope|KunKun|ColorIjo|J1ngga|Grandong.A|Grandong.B|Troj.Starter.Y|Troj.Ramnit.A.46|Malingsia.C|Malingsia.B|Malingsia.A|Playboy.B|Playboy.A|Shiren Sungkar|Ijab Qobul|Cinta Laura|ABGila|ZnWarnet|Jim|ojanBLANK|Trojan.cript|Conficker.A|Conficker.B|Conficker.G|VB.Bodoh|VB.Rnd|Troj.Random|Troj.runer"

Public Const SalityA As String = "E000"
Public Const SalityB As String = "F000"

Public IDb(1 To 500) As String
Public ttlDB As Long
Public pescraM As String
Public PEtemp As String
Public RamnitSrc As String
Public REGrun As String
Public REGhiden As String
