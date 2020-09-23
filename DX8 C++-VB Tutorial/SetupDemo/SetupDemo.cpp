// SetupDemo.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"         //Included automatically
#include <Math.h>           //Same as the header file
#include "SetupDemo.h"      //Include or header file 

bool _stdcall DllMain(HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH: 
	case DLL_THREAD_DETACH: 
case DLL_PROCESS_DETACH: break;
}

return true;
}

_declspec(dllexport) bool _stdcall FlipMeshNormals(D3DVERTEX *Vertices, long NumVertices)
{
	try{
		for(long i = 0; i < NumVertices; i++)
		{
			Vertices->nX = -Vertices->nX;
			Vertices->nY = -Vertices->nY;
			Vertices->nZ = -Vertices->nZ;
			Vertices++;
		}
		return true;
	}catch(...){
		return false;
	}
}

_declspec(dllexport) float _stdcall Distance2D(D3DVECTOR2& Vector1, D3DVECTOR2& Vector2)
{
	return sqrtf((Vector2.X - Vector1.X) * (Vector2.X - Vector1.X) + 
                 (Vector2.Y - Vector1.Y) * (Vector2.Y - Vector1.Y));
}

