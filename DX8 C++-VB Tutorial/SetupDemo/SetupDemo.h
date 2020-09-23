
//Included Header Files
#include <Math.h>

//Type Definition Structures
typedef struct D3DVECTOR2
{
	float X;
	float Y;
} _D3DVECTOR2;

typedef struct D3DVECTOR
{
	float X;
	float Y;
	float Z;
} _D3DVECTOR;

typedef struct D3DVECTOR4
{
	float X;
	float Y;
	float Z;
	float W;
} _D3DVECTOR4;

typedef struct D3DVERTEX
{
	float X;
	float Y;
	float Z;
	float nX;
	float nY;
	float nZ;
	float tU;
	float tV;
} _D3DVERTEX;

typedef struct D3DVERTEX2
{
	float X;
	float Y;
	float Z;
	float nX;
	float nY;
	float nZ;
	float tU1;
	float tV1;
	float tU2;
	float tV2;
} _D3DVERTEX2;

typedef struct D3DLVERTEX
{
	float X;
	float Y;
	float Z;
	float tU;
	float tV;
	long Color;
	long Specular;
} _D3DLVERTEX;

typedef struct D3DLVERTEX2
{
	float X;
	float Y;
	float Z;
	float tU1;
	float tV1;
	float tU2;
	float tV2;
	long Color;
	long Specular;
} _D3DLVERTEX2;

typedef struct D3DTLVERTEX
{
	float X;
	float Y;
	float Z;
	float tU;
	float tV;
	float RHW;
	long Color;
	long Specular;
} _D3DTLVERTEX;

typedef struct D3DTLVERTEX2
{
	float X;
	float Y;
	float Z;
	float tU1;
	float tV1;
	float tU2;
	float tV2;
	float RHW;
	long Color;
	long Specular;
} _D3DTLVERTEX2;

//Prototypes
_declspec(dllexport) bool _stdcall FlipMeshNormals(D3DVERTEX *Vertices, long NumVertices);
_declspec(dllexport) float _stdcall Distance2D(D3DVECTOR2& Vector1, D3DVECTOR2& Vector2);