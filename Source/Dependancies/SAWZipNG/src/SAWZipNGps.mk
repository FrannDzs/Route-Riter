
SAWZipNGps.dll: dlldata.obj SAWZipNG_p.obj SAWZipNG_i.obj
	link /dll /out:SAWZipNGps.dll /def:SAWZipNGps.def /entry:DllMain dlldata.obj SAWZipNG_p.obj SAWZipNG_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del SAWZipNGps.dll
	@del SAWZipNGps.lib
	@del SAWZipNGps.exp
	@del dlldata.obj
	@del SAWZipNG_p.obj
	@del SAWZipNG_i.obj
