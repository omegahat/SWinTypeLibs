/*
  See http://spec.winprog.org/typeinf2/
  for a useful tutorial on type libraries.
*/
#include <ole2.h>
#include <oleauto.h>

#include <objbase.h>

#include <stdio.h>

#ifdef ERROR
#undef ERROR
#endif

extern "C" {
#include <Rdefines.h>
#include <R_ext/Rdynload.h>

#include "RError.h"
}

extern "C" {
  __declspec(dllexport) SEXP R_loadTypeLib(SEXP fileName);
  __declspec(dllexport) SEXP R_getTypeLibInfoCount(SEXP obj);
  __declspec(dllexport) void R_SWinTypeLibs_init(DllInfo *info);
  __declspec(dllexport) SEXP R_getTypeInfoTypes(SEXP obj, SEXP which);
  __declspec(dllexport) SEXP R_getTypeLibDocumentation(SEXP obj, SEXP which, SEXP isTypeLib);

  __declspec(dllexport) SEXP R_getTypeLibInfoEntry(SEXP s_info, SEXP which);

  __declspec(dllexport) SEXP R_getDCOMInfoEntry(SEXP obj, SEXP which);

  __declspec(dllexport) SEXP R_getTypeLibNames(SEXP obj);

  __declspec(dllexport) SEXP R_getEnums(SEXP s_info);
  __declspec(dllexport) SEXP R_getCVars(SEXP s_info);
  __declspec(dllexport) SEXP R_getCFuncs(SEXP s_info);
  __declspec(dllexport) SEXP R_getCoClass(SEXP s_info);
  __declspec(dllexport) SEXP R_getInterfaces(SEXP s_info);

  __declspec(dllexport) SEXP R_getInvokeEnum();

  __declspec(dllexport) SEXP R_getTypeInfoDocumentation(SEXP s_info, SEXP idx);
  __declspec(dllexport) SEXP R_getTypeInfoCount(SEXP obj);

  __declspec(dllexport) SEXP R_GetNames(SEXP s_info, SEXP sid);

  __declspec(dllexport) SEXP R_GetGuid(SEXP s_info);

  __declspec(dllexport) SEXP R_getDCOMInfoCount(SEXP obj);

  __declspec(dllexport) SEXP R_getTypeInfoHRefType(SEXP obj, SEXP reftype);

  __declspec(dllexport) SEXP R_GetIDsOfNames(SEXP s_info, SEXP s_names, SEXP isCOMObject);

  __declspec(dllexport) SEXP R_GetIDsOfFuncNames(SEXP s_info, SEXP s_mid);

  __declspec(dllexport) SEXP R_getRefTypeInfo(SEXP s_info, SEXP which);

  __declspec(dllexport) SEXP R_getTypeLibUUIDs(SEXP obj);

  __declspec(dllexport) SEXP R_GetTypeInfoLibrary(SEXP info);

  __declspec(dllexport) SEXP R_getAlias(SEXP s_info);

  __declspec(dllexport) SEXP R_getLibraryAttr(SEXP s_lib);

  __declspec(dllexport) SEXP R_loadRegTypeLib(SEXP guid, SEXP version, SEXP lcid);

  __declspec(dllexport) SEXP R_QueryPathOfRegTypeLib(SEXP guid, SEXP version, SEXP lcid);
}

/* Why not use the TKIND enumerations here ? */

typedef enum {FUNCTIONS, VARIABLES, COCLASS, ENUMERATION, INTERFACES, ALIAS, MEANINGLESS} R_TypeInfoKind;

SEXP R_getTypeInfoEls(SEXP s_info, R_TypeInfoKind type);

void *getRReference(SEXP s);
void GetScodeString(HRESULT hr, LPTSTR buf, int bufSize);
SEXP R_createRTypeRefObject(void *ref, const char *className, R_CFinalizer_t finalizer);
SEXP R_newFunctionObject(FUNCDESC *);
SEXP R_newVariableObject();
SEXP R_newObject(const char *className);
const char *getPrimitiveTypeName(int which);
SEXP R_describeType(TYPEDESC *desc, ITypeInfo *);
SEXP R_getCFuncDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *);
SEXP R_getCVarDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *);
SEXP R_getEnumDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *);
SEXP R_getCoClassDesc(TYPEATTR *attr, int which, ITypeInfo *info, ITypeInfo **);
SEXP R_describeParameter(ELEMDESC *edesc, ITypeInfo *info);

SEXP setTypeInfoTypeSlot(SEXP obj, ITypeInfo *type);

static SEXP GetNames(ITypeInfo *info, MEMBERID id, UINT count, UINT offset);

const char * const getTypeKindName(TYPEKIND val);

static const char *FunctionKindName(FUNCKIND kind);

static SEXP R_ITypeLibSym;

const char *FromBstr(BSTR str);

__declspec(dllexport) 
void R_SWinTypeLibs_init(DllInfo *info)
{
  R_ITypeLibSym = Rf_install("ITypeLib");
}


void
COMError(HRESULT hr)
{
    TCHAR buf[512];
    SEXP e;
    GetScodeString(hr, buf, sizeof(buf)/sizeof(buf[0]));
    /*
    PROBLEM buf
    ERROR;
    */

    PROTECT(e = allocVector(LANGSXP, 3));
    SETCAR(e, Rf_install("COMStop"));
    SETCAR(CDR(e), mkString(buf));
    SETCAR(CDR(CDR(e)), ScalarInteger(hr));
    Rf_eval(e, R_GlobalEnv);
    UNPROTECT(1); /* Won't come back to here. */

}


void
R_typelib_finalizer(SEXP s)
{
 ITypeLib *lib;
 lib = (ITypeLib *) R_ExternalPtrAddr(s);

 if(lib) {
   R_ClearExternalPtr(s);
   lib->Release();  /*XXX ? */
 }
}



SEXP
R_createRTypeLib(void *ref)
{
  return(R_createRTypeRefObject(ref, "ITypeLib", R_typelib_finalizer));
}

SEXP
R_createRTypeRefObject(void *ref, const char *className, R_CFinalizer_t finalizer)
{
  SEXP ans, obj, klass;

  PROTECT(ans = R_MakeExternalPtr((void*) ref, Rf_install(className), R_NilValue));
  if(finalizer)
     R_RegisterCFinalizer(ans, finalizer);

  klass = MAKE_CLASS((char *) className);
  PROTECT(obj = NEW(klass));
  SET_SLOT(obj, Rf_install("ref"), ans);

  UNPROTECT(2);

  return(obj);
}

BSTR
AsBstr(const char *str)
{
  BSTR ans = NULL;

  if(!str)
    return(NULL);

  int size = strlen(str);
  int wideSize = 2 * size;

  LPOLESTR wstr = (LPWSTR) S_alloc(wideSize, sizeof(OLECHAR)); //XXX sizeof(char));
  if(!wstr) {
    PROBLEM "Can't allocate space for DCOM string (BSTR) of size %d", wideSize
      ERROR;
  }

  int hr = MultiByteToWideChar(CP_ACP, 0, str, size, wstr, wideSize);
  if(hr == 0) {
    PROBLEM "Can't create BSTR"
    ERROR;
  }
  ans = SysAllocStringLen(wstr, size);

  return(ans);
}

__declspec(dllexport) SEXP 
R_loadTypeLib(SEXP fileName)
{
  ITypeLib *type;
  HRESULT hr;
  BSTR str;

  str = AsBstr(CHAR(STRING_ELT(fileName, 0)));
  hr = LoadTypeLib(str, &type);
  SysFreeString(str);

  if(FAILED(hr)) 
    COMError(hr);

  type->AddRef();

  return(R_createRTypeLib(type));
}

extern int R_getCLSIDFromString(SEXP className, CLSID *classId);

__declspec(dllexport) SEXP 
R_loadRegTypeLib(SEXP guid, SEXP version, SEXP lcid)
{
  ITypeLib *type;
  HRESULT hr;
  CLSID g;
  LCID locale = REAL(lcid)[0];

  hr = R_getCLSIDFromString(guid, &g);
  hr = LoadRegTypeLib(g, INTEGER(version)[0], INTEGER(version)[1], locale, &type);
  if(FAILED(hr)) 
    COMError(hr);

  type->AddRef();

  return(R_createRTypeLib(type));
}



__declspec(dllexport) SEXP 
R_QueryPathOfRegTypeLib(SEXP guid, SEXP version, SEXP lcid)
{
  HRESULT hr;
  CLSID g;
  BSTR str;
  SEXP ans = R_NilValue;

  LPOLESTR wstr = (LPWSTR) S_alloc(800 * 2, sizeof(OLECHAR)); 
  str = SysAllocStringLen(wstr, 800);

  hr = R_getCLSIDFromString(guid, &g);
  hr = QueryPathOfRegTypeLib(g, INTEGER(version)[0], INTEGER(version)[1], 0, &str);
  if(FAILED(hr)) 
    COMError(hr);

  ans = Rf_mkString(FromBstr(str));
  SysFreeString(str);

  return(ans);
}




void *
getRReference(SEXP s)
{
  void *ptr;
  SEXP el;
  
  if(TYPEOF(s) == EXTPTRSXP) {
     PROBLEM "getRRreference called with an external pointer"
       ERROR;
  }

  el = GET_SLOT(s, Rf_install("ref"));
  ptr = R_ExternalPtrAddr(el);
  if(!ptr)
    PROBLEM "NULL value passed to C in an R external pointer. (Was this saved in a previous session?)"
      ERROR;

  return(ptr);
}


__declspec(dllexport) SEXP 
R_getTypeLibInfoCount(SEXP obj)
{
 ITypeLib* disp;
 SEXP ans;
 unsigned int val = 0;

 disp = (ITypeLib *) getRReference(obj);

 val = disp->GetTypeInfoCount();

 ans = NEW_INTEGER(1);
 INTEGER_DATA(ans)[0] = val;

 return(ans);
}


__declspec(dllexport) SEXP
R_getTypeInfoTypes(SEXP obj, SEXP which)
{
  int i = 0;
  int n = Rf_length(which);
  ITypeLib *lib = (ITypeLib *) getRReference(obj);
  SEXP ans;
  TYPEKIND type;

  PROTECT(ans = NEW_INTEGER(n));
  for(i = 0; i < n; i++) {
    lib->GetTypeInfoType((unsigned int) INTEGER_DATA(which)[i], &type);
    INTEGER_DATA(ans)[i] = type;
  }
  UNPROTECT(1);

  return(ans);
}



__declspec(dllexport) SEXP
R_getTypeLibNames(SEXP obj)
{
  int i = 0;
  int n;
  ITypeLib *lib = (ITypeLib *) getRReference(obj);
  SEXP ans;

  n = lib->GetTypeInfoCount();
  PROTECT(ans = NEW_CHARACTER(n));
  for(i = 0; i < n; i++) {
    BSTR name;
    lib->GetDocumentation((unsigned int) i, &name, NULL, NULL, NULL);
    SET_STRING_ELT(ans, i, COPY_TO_USER_STRING(FromBstr(name)));
  }
  UNPROTECT(1);

  return(ans);
}

__declspec(dllexport) SEXP
R_getTypeLibUUIDs(SEXP obj)
{
  int i = 0;
  int n;
  ITypeLib *lib = (ITypeLib *) getRReference(obj);
  SEXP ans;

  n = lib->GetTypeInfoCount();
  PROTECT(ans = NEW_CHARACTER(n));
  for(i = 0; i < n; i++) {
    BSTR ostr;
    HRESULT hr;
    ITypeInfo *type;
    TYPEATTR *typeAttr;
    hr = lib->GetTypeInfo(i, &type); //XXX do we need to release this
    hr = type->GetTypeAttr(&typeAttr);
    if(FAILED(hr) ||!type)
      COMError(hr);

    hr = StringFromCLSID(typeAttr->guid, &ostr);
    if(hr == S_OK) {
     SET_STRING_ELT(ans, i, COPY_TO_USER_STRING(FromBstr(ostr)));
     SysFreeString(ostr);
    }
    type->ReleaseTypeAttr(typeAttr);
    type->Release();
  }
  UNPROTECT(1);

  return(ans);
}


static const char* emptyString = "";

const char *
FromBstr(BSTR str)
{
  char *ptr;
  DWORD len;

  if(!str)
    return(emptyString);

  len = wcslen(str);
  if(len == 0)
    return(emptyString);

  ptr = (char *) S_alloc(len + 1, sizeof(char));
  if(!ptr) { /* S_alloc handles errors for us.*/
    PROBLEM "Can't allocate space for converting DCOM string to R (length %d)", (int) (len + 1)
     ERROR;
  }
  ptr[len] = '\0';

  DWORD ok = WideCharToMultiByte(CP_ACP, 0, str, len, ptr, len, NULL, NULL);
  if(ok == 0) {
    PROBLEM "Converting FromBstr"
      ERROR;
    return(emptyString);
  }

  return(ptr);
}

SEXP
R_getTypeInfoHRefType(SEXP obj, SEXP reftype)
{
  ITypeInfo *info, *type;
  SEXP tmp;
  HRESULT hr;
  HREFTYPE href;

  info = (ITypeInfo *) getRReference(obj);
  href = (HREFTYPE) NUMERIC_DATA(reftype)[0];
  hr = info->GetRefTypeInfo(href, &type);

  if(FAILED(hr) || !type) {
    PROBLEM "Can't get ITypeInfo from hreftype."
    ERROR;
  }

  PROTECT(tmp = R_createRTypeRefObject(type, "ITypeInfo", NULL));
  setTypeInfoTypeSlot(tmp, type);
  UNPROTECT(1);

  return(tmp);
}

__declspec(dllexport) SEXP
R_getTypeLibDocumentation(SEXP obj, SEXP which, SEXP isTypeLib)
{
  SEXP ans;
  BSTR name, doc, helpFile;
  HRESULT hr;
  ITypeInfo *info;
  ITypeLib *lib;

  if(LOGICAL(isTypeLib)[0]) {
    lib = (ITypeLib *) getRReference(obj);
    hr = lib->GetDocumentation(INTEGER(which)[0], &name, &doc, NULL, &helpFile);
  } else {
    info = (ITypeInfo *) getRReference(obj);
    hr = info->GetDocumentation(INTEGER(which)[0], &name, &doc, NULL, &helpFile);
  }

  if(FAILED(hr)) 
    COMError(hr);

  PROTECT(ans = NEW_CHARACTER(3));
  SET_STRING_ELT(ans, 0, COPY_TO_USER_STRING(FromBstr(name)));
  SET_STRING_ELT(ans, 1, COPY_TO_USER_STRING(FromBstr(doc)));
  SET_STRING_ELT(ans, 2, COPY_TO_USER_STRING(FromBstr(helpFile)));


  SEXP names;
  PROTECT(names = NEW_CHARACTER(3));
  const char * const elNames[] = {"name", "documentation", "helpFile"};
  for(int i = 0; i < 3; i++) {
    SET_STRING_ELT(names, i, COPY_TO_USER_STRING(elNames[i]));
  }
  SET_NAMES(ans, names);
  UNPROTECT(2);

  SysFreeString(name);
  SysFreeString(doc);
  SysFreeString(helpFile);

  return(ans);
}

SEXP 
setTypeInfoTypeSlot(SEXP obj, ITypeInfo *type)
{
  SEXP ans, names;
  TYPEATTR *typeAttr = NULL;
  char *str;

  if(!type)
    return(obj);

  type->GetTypeAttr(&typeAttr);

  if(!typeAttr)
    return(obj);

  PROTECT(obj);

  PROTECT(ans = NEW_INTEGER(1));
  PROTECT(names = NEW_CHARACTER(1));
  
   INTEGER_DATA(ans)[0] = typeAttr->typekind;
   str = (char *) getTypeKindName(typeAttr->typekind);
   SET_STRING_ELT(names, 0, COPY_TO_USER_STRING(str));
   SET_NAMES(ans, names);

   SET_SLOT(obj, Rf_install("type"), ans);
  UNPROTECT(2);

  
   BSTR ostr;
   HRESULT hr = StringFromCLSID(typeAttr->guid, &ostr);


   
   if(hr == S_OK) {
     PROTECT(names = NEW_CHARACTER(1));
     SET_STRING_ELT(names, 0, COPY_TO_USER_STRING(FromBstr(ostr)));
     SET_SLOT(obj, Rf_install("guid"), names);

     // original SysFreeString(ostr);
     // google windows StringFromCLSID and see how to free -> CoTaskMemFree()
     CoTaskMemFree(ostr);
     UNPROTECT(1);
   }

  UNPROTECT(1);
  type->ReleaseTypeAttr(typeAttr);
  return(obj);
}

typedef struct _ITypeInfoClassNames {
  const char *className;
  int /*tagTYPEKIND*/ type;
} ITypeInfoClassMap;

const ITypeInfoClassMap ITypeInfoClassNames[] = {
  {"ITypeInfoEnum", TKIND_ENUM},
  {"ITypeInfoRecord", TKIND_RECORD},
  {"ITypeInfoModule", TKIND_MODULE},
  {"ITypeInfoInterface", TKIND_INTERFACE},
  {"ITypeInfoDispatch", TKIND_DISPATCH},
  {"ITypeInfoCoClass", TKIND_COCLASS},
  {"ITypeInfoAlias", TKIND_ALIAS},
  {"ITypeInfoUnion", TKIND_UNION}
};


static const char *
getTypeInfoClassName(ITypeInfo *type)
{
  TYPEATTR *typeAttr;
  static const char *str = "ITypeInfo";
  unsigned int i;

  type->GetTypeAttr(&typeAttr);
  if(!typeAttr)
    return("ITypeInfo");

  for(i = 0; i < sizeof(ITypeInfoClassNames)/sizeof(ITypeInfoClassNames[0]); i++) {
	if(typeAttr->typekind == ITypeInfoClassNames[i].type) {
	  str = ITypeInfoClassNames[i].className;
	  break;
	}
  }

  type->ReleaseTypeAttr(typeAttr);
  return(str);
}

__declspec(dllexport) 
SEXP
R_getTypeLibInfoEntry(SEXP obj, SEXP which)
{
  ITypeLib *lib = (ITypeLib *) getRReference(obj);
  SEXP ans = R_NilValue, tmp;
  ITypeInfo *type;
  HRESULT hr;
  int i, n = Rf_length(which);
  int useOwnIndex = 0;

  /* This doesn't happen now as we compute the index vector using 1:length(lib) */
  if(n == 0) {
    n = lib->GetTypeInfoCount();
    useOwnIndex = 1;
  }

  PROTECT(ans = NEW_LIST(n));
  for(i = 0; i < n; i++) {
    hr = lib->GetTypeInfo(useOwnIndex ? i : INTEGER(which)[i], &type);

    if(!FAILED(hr) && type) {
      PROTECT(tmp = R_createRTypeRefObject(type, getTypeInfoClassName(type), NULL));
      setTypeInfoTypeSlot(tmp, type);
      SET_VECTOR_ELT(ans, i, tmp);
      UNPROTECT(1);
      type->Release();
    } else {
      COMError(hr);
    }
  }
  UNPROTECT(1);
  return(ans);
}


SEXP
R_GetTypeInfoLibrary(SEXP sinfo)
{
  ITypeLib *lib;
  ITypeInfo *info;
  HRESULT hr;

  info = (ITypeInfo *) getRReference(sinfo);
  hr = info->GetContainingTypeLib(&lib, NULL);
  if(FAILED(hr)) {
    PROBLEM "Can't get ITypeLib from ITypeInfo."
    ERROR;
  }

  return(R_createRTypeRefObject(lib, "IContainingTypeLib", R_typelib_finalizer));
}


SEXP
R_getLibraryAttr(SEXP s_lib)
{
  TLIBATTR *attr = NULL;
  int nprotect = 1;
  ITypeLib *lib = (ITypeLib *) getRReference(s_lib);
  HRESULT hr;
  SEXP ans = R_NilValue, tmp;

  hr = lib->GetLibAttr(&attr);
  if(hr != S_OK) {
    return R_NilValue;
  }
 
  PROTECT(ans = NEW_LIST(6));

  LPOLESTR ostr;
  hr = StringFromCLSID(attr->guid, &ostr);
  if(hr == S_OK) 
    SET_VECTOR_ELT(ans, 0, Rf_mkString(FromBstr(ostr)));

  SET_VECTOR_ELT(ans, 1, Rf_ScalarReal(attr->lcid));

  PROTECT(tmp = Rf_ScalarInteger(attr->syskind));nprotect++;
  SET_NAMES(tmp, Rf_mkString(attr->syskind == SYS_WIN16 ? 
			     "WIN16" : attr->syskind == SYS_WIN32 ? "WIN32" : "MAC"));
  SET_VECTOR_ELT(ans, 2, tmp);
  SET_VECTOR_ELT(ans, 3, Rf_ScalarInteger(attr->wMajorVerNum));
  SET_VECTOR_ELT(ans, 4, Rf_ScalarInteger(attr->wMinorVerNum));
  SET_VECTOR_ELT(ans, 5, Rf_ScalarInteger(attr->wLibFlags));


  lib->ReleaseTLibAttr(attr);

  UNPROTECT(nprotect);

  return(ans);  
}


__declspec(dllexport)
SEXP
R_getDCOMInfoCount(SEXP obj)
{
  UINT n;
  HRESULT hr;
  SEXP ans;
  IDispatch *ref = (IDispatch *) R_ExternalPtrAddr(obj);

  if(!ref) {
    PROBLEM "Null COM object"
    ERROR;
  }
  hr = ref->GetTypeInfoCount(&n);
  if(FAILED(hr))
      COMError(hr);
  ans = NEW_INTEGER(1);
  INTEGER_DATA(ans)[0] = n;
  return(ans);
}

__declspec(dllexport)
SEXP
R_getDCOMInfoEntry(SEXP obj, SEXP which)
{
  IDispatch *ref = (IDispatch *) R_ExternalPtrAddr(obj);
  SEXP ans = R_NilValue, tmp;
  ITypeInfo *type;
  HRESULT hr;
  int useOwnIndex = 0;
  UINT n, i;
  if(!ref) {
    PROBLEM "Null COM object"
    ERROR;
  }

  n = Rf_length(which);
  /* This doesn't happen now as we compute the index vector using 1:length(lib) */
  if(n == 0) {
    hr = ref->GetTypeInfoCount(&n);
    if(FAILED(hr))
      return(NULL_USER_OBJECT);
  }

  PROTECT(ans = NEW_LIST(n));
  for(i = 0; i < n; i++) {
    hr = ref->GetTypeInfo(useOwnIndex ? i : INTEGER(which)[i], 0, &type);
    if(!FAILED(hr) && type) {
      PROTECT(tmp = R_createRTypeRefObject(type, "ITypeInfo", NULL));
      setTypeInfoTypeSlot(tmp, type);
      SET_VECTOR_ELT(ans, i, tmp);
      UNPROTECT(1);
    } else {
      COMError(hr);
    }
  }
  
  UNPROTECT(1);
  return(ans);
}

void
R_releaseTypeAttr(SEXP s)
{
#if 0
 TYPEATTR *ptr =  R_ExternalPtrAddr(s);
#endif
}

SEXP
R_getTypeInfoAttr(SEXP s_info)
{
  ITypeInfo *type = (ITypeInfo *) getRReference(s_info);
  SEXP ans = R_NilValue;
  HRESULT hr;
  TYPEATTR *typeAttr;

  hr = type->GetTypeAttr(&typeAttr);
  if(!FAILED(hr) && typeAttr) {
    ans = R_createRTypeRefObject(typeAttr, "TYPEATTR", R_releaseTypeAttr);
  } else {
    COMError(hr);
  }

  return(ans);
}

__declspec(dllexport) 
SEXP
R_getCFuncs(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, FUNCTIONS));
}

__declspec(dllexport) 
SEXP
R_getInterfaces(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, INTERFACES));
}


__declspec(dllexport) 
SEXP
R_getCVars(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, VARIABLES));
}

__declspec(dllexport) 
SEXP
R_getEnums(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, ENUMERATION));
}

__declspec(dllexport) 

SEXP
R_getCoClass(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, COCLASS));
}

SEXP
R_getAlias(SEXP s_info)
{
  return(R_getTypeInfoEls(s_info, ALIAS));
}

#if 0
__declspec(dllexport) 
SEXP
R_getDesc(SEXP s_info)
{
  ITypeInfo *type = (ITypeInfo *) getRReference(s_info);
  hr = ref->GetTypeInfo(INTEGER(which)[0], 0, &type);
  return(R_getTypeInfoEls(s_info, ));
}
#endif


__declspec(dllexport)
SEXP
R_getTypeInfoDocumentation(SEXP s_info, SEXP idx)
{
  ITypeInfo *type = (ITypeInfo *) getRReference(s_info);
  HRESULT hr;
  BSTR name, desc;
  int n, i;
  bool getAll = INTEGER_DATA(idx)[0] < 0;

  if(getAll) {
    n = 0;
    while((hr = type->GetDocumentation(n, &name, &desc, NULL, NULL)) == ERROR_SUCCESS)
      n++;
  } else
    n = 1;

  SEXP names, ans = R_NilValue;
  PROTECT(ans = NEW_CHARACTER(n));
  PROTECT(names = NEW_CHARACTER(n));
  for(i = 0 ; i < n ; i++) {
    hr = type->GetDocumentation(getAll ? i : INTEGER_DATA(idx)[0], &name, &desc, NULL, NULL);
    if(hr == ERROR_SUCCESS) {
      SET_STRING_ELT(ans, i, COPY_TO_USER_STRING(FromBstr(desc)));
      SET_STRING_ELT(names, i, COPY_TO_USER_STRING(FromBstr(name)));
      SysFreeString(name);
      SysFreeString(desc);
    }
  }
  SET_NAMES(ans, names);
  UNPROTECT(2);
  return(ans);
}


#if 0
/* ITypeInfo Doesn't have a GetTypeInfoCount*/
__declspec(dllexport) SEXP 
R_getTypeInfoCount(SEXP s_info)
{
 SEXP ans;
 unsigned int val = 0;
 ITypeInfo *type = (ITypeInfo *) getRReference(s_info);

 val = type->GetTypeInfoCount();

 ans = NEW_INTEGER(1);
 INTEGER(ans)[0] = val;
 return(ans);
}
#endif


SEXP
R_getTypeInfoEls(SEXP s_info, R_TypeInfoKind target)
{
  ITypeInfo *type = (ITypeInfo *) getRReference(s_info);
  SEXP ans;
  int i, n = 0;
  TYPEATTR *typeAttr;
  MEMBERID memid;

  type->GetTypeAttr(&typeAttr);
  if(!typeAttr) {
    PROBLEM "Can't get Type Attribute from ITypeInfo"
    ERROR;
  }
 
  switch(target) {
    case FUNCTIONS:
     n = typeAttr->cFuncs; 
     break;
    case COCLASS:
    case INTERFACES:
      n = typeAttr->cImplTypes; // + 1 is if this is a dual interface
      //and then we want to get the -1
     break;
    case VARIABLES:
    case ENUMERATION:
     n = typeAttr->cVars; 
     break;
    case ALIAS:
      {
	ans = R_describeType(&typeAttr->tdescAlias, type);
        type->ReleaseTypeAttr(typeAttr);
        return(ans);
      }
     break;
    default:
     PROBLEM "Unsupported type"
      ERROR;
    }

  if(n == 0)
    return(R_NilValue);

  PROTECT(ans = NEW_LIST(n));
#if 0
  PROTECT(names = NEW_CHARACTER(n));
#endif
  for(i = 0; i < n; i++) {
    SEXP el;
    memid = -1;
    switch(target) {
     case FUNCTIONS:
      el = R_getCFuncDesc(typeAttr, i, type, &memid);
      break;
    case COCLASS:
    case INTERFACES:
      {
       ITypeInfo *hinfo = NULL;
       el = R_getCoClassDesc(typeAttr, i, type, &hinfo);
       BSTR name;
       hinfo->GetDocumentation(0, &name, NULL, NULL, NULL);
#if 0
       SET_STRING_ELT(names, i, COPY_TO_USER_STRING(FromBstr(name)));
#endif
       SysFreeString(name);
      }
      break;
    case VARIABLES:
      el = R_getCVarDesc(typeAttr, i, type, &memid);
      break;
     case ENUMERATION:
      el = R_getEnumDesc(typeAttr, i, type, &memid);
      break;
     default:
      el = NULL;
      break;
    }
    if(el) {
      SEXP tmp;
      PROTECT(tmp = NEW_LOGICAL(1));
      if(typeAttr->wTypeFlags & TYPEFLAG_FHIDDEN) {
	LOGICAL_DATA(tmp)[0] = TRUE;
      }
      SET_SLOT(el, Rf_install("hidden"), tmp);
      UNPROTECT(1);

      SET_VECTOR_ELT(ans, i, el);

#if 0
      if(memid != -1) {
       BSTR name;
       type->GetDocumentation(memid, &name, NULL, NULL, NULL);
       SET_STRING_ELT(names, i, COPY_TO_USER_STRING(FromBstr(name)));
       SysFreeString(name);
      }
#endif
    }
  }
#if 0
  SET_NAMES(ans, names);
#endif
  UNPROTECT(1); //XXX make this 2 if we turn on the names
  type->ReleaseTypeAttr(typeAttr);
  return(ans);
}

SEXP
R_getCoClassDesc(TYPEATTR *attr, int which, ITypeInfo *info, ITypeInfo **refInfo)
{
  HREFTYPE refType;
  ITypeInfo *hinfo;
  SEXP ans;

  info->GetRefTypeOfImplType(which, &refType);
  info->GetRefTypeInfo(refType, &hinfo);
  *refInfo = hinfo;

  PROTECT(ans = R_createRTypeRefObject(hinfo, getTypeInfoClassName(hinfo), NULL));
  setTypeInfoTypeSlot(ans, hinfo); /* TKIND_DISPATCH; */
  UNPROTECT(1);
  
  return(ans);
}

SEXP
R_getRefTypeInfo(SEXP s_info, SEXP href)
{
  ITypeInfo *info = (ITypeInfo *) getRReference(s_info);
  HREFTYPE refType;
  HRESULT hr;
  ITypeInfo *hinfo;

  refType = (HREFTYPE) REAL(href)[0];

  hr = info->GetRefTypeInfo(refType, &hinfo);

  if(hr != S_OK) {
      PROBLEM  "GetRefTypeInfo() error : %ld", hr
      ERROR;
  }
  
  SEXP ans;
  PROTECT(ans = R_createRTypeRefObject(hinfo, "ITypeInfo", NULL));
  setTypeInfoTypeSlot(ans, hinfo); /* TKIND_DISPATCH; */
  UNPROTECT(1);
  return(ans);
}

/* No function to call  this yet. */
SEXP
R_getRefTypeOfImplInfo(SEXP s_info, SEXP which)
{
  ITypeInfo *info = (ITypeInfo *) getRReference(s_info);
  HREFTYPE refType;
  ITypeInfo *hinfo;
  HRESULT hr;

  hr = info->GetRefTypeOfImplType(INTEGER_DATA(which)[0], &refType);
  if(hr)
    return(R_NilValue);
  hr = info->GetRefTypeInfo(refType, &hinfo);
  if(hr)
    return(R_NilValue);

  SEXP ans;
  PROTECT(ans = R_createRTypeRefObject(hinfo, "ITypeInfo", NULL));
  setTypeInfoTypeSlot(ans, hinfo); /* TKIND_DISPATCH; */
  UNPROTECT(1);
  return(ans);
}

SEXP 
R_getEnumDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *memid)
{
  SEXP ans, names;
  VARDESC *desc;
  HRESULT hr;
  BSTR name;

  hr = info->GetVarDesc(which, &desc);

  if(FAILED(hr)) {
     return(R_NilValue);
  }

  PROTECT(ans = NEW_INTEGER(1));
  INTEGER_DATA(ans)[0] = V_I4(desc->lpvarValue);

  PROTECT(names = NEW_CHARACTER(1));
  *memid = desc->memid;
  info->GetDocumentation(desc->memid, &name, NULL, NULL, NULL);
  SET_STRING_ELT(names, 0, COPY_TO_USER_STRING(FromBstr(name)));
  SysFreeString(name);

  SET_NAMES(ans, names);
  UNPROTECT(2);

  return(ans);
}



SEXP 
R_getCVarDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *memid)
{
  SEXP el, ans;
  VARDESC *desc;
  HRESULT hr;
  BSTR name;
  ELEMDESC *edesc;

  hr = info->GetVarDesc(which, &desc);

  if(FAILED(hr)) {
     return(R_NilValue);
  }

  info->GetDocumentation(desc->memid, &name, NULL, NULL, NULL);
  *memid = desc->memid;

  PROTECT(ans = R_newVariableObject());
  PROTECT(el = NEW_CHARACTER(1));
  SET_STRING_ELT(el, 0, COPY_TO_USER_STRING(FromBstr(name)));
  SET_SLOT(ans, Rf_install("name"), el);
  UNPROTECT(1);

  edesc = &desc->elemdescVar;
  SET_SLOT(ans, Rf_install("type"), R_describeType(&(edesc->tdesc), info));

  UNPROTECT(1); /* ans */

  SysFreeString(name);
  info->ReleaseVarDesc(desc);

  return(ans);
}

SEXP
R_setMemId(MEMBERID id, SEXP obj)
{
  SEXP tmp;
  PROTECT(tmp = NEW_INTEGER(1));
  INTEGER_DATA(tmp)[0] = id;
  SET_SLOT(obj, Rf_install("memid"), tmp);
  UNPROTECT(1);
  return(obj);
}


/*
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfsystemruntimeinteropservicesfuncdescclasslprgscodetopic.asp
*/
SEXP
R_getCFuncDesc(TYPEATTR *attr, int which, ITypeInfo *info, MEMBERID *memid)
{
  SEXP el, ans;
  FUNCDESC *desc;
  HRESULT hr;
  BSTR name;
  int i;

  hr = info->GetFuncDesc(which, &desc);

  if(FAILED(hr)) {
     return(R_NilValue);
  }

  info->GetDocumentation(desc->memid, &name, NULL, NULL, NULL);
  *memid = desc->memid;

  PROTECT(ans = R_newFunctionObject(desc));
  R_setMemId(desc->memid, ans);
  el = GET_SLOT(ans, Rf_install("name"));
  SET_STRING_ELT(el, 0, COPY_TO_USER_STRING(FromBstr(name)));
  SysFreeString(name);

  if(desc->cParams) {
    PROTECT(el = NEW_LIST(desc->cParams));
     for(i = 0; i < desc->cParams ; i++) {
        SET_VECTOR_ELT(el, i, R_describeParameter(&desc->lprgelemdescParam[i], info));
     }
     SET_NAMES(el, GetNames(info, desc->memid, desc->cParams+1, 1));
     SET_SLOT(ans, Rf_install("parameters"), el);
     UNPROTECT(1);
  }

  /* Potentially we have multiple permitted types in cScodes */
  SET_SLOT(ans, Rf_install("returnType"), R_describeType(&desc->elemdescFunc.tdesc, info));

  el = GET_SLOT(ans, Rf_install("invokeType"));
  INTEGER(el)[0] = desc->invkind;
  //XXX get the symbolic name of this value.  virtual, pure-virtual, nonvirtual, static, dispatch.
  // see below
  SEXP tmp, names;
  PROTECT(tmp = NEW_INTEGER(1));
  PROTECT(names = NEW_CHARACTER(1));
  INTEGER(tmp)[0] = desc->funckind;
  SET_STRING_ELT(names, 0, COPY_TO_USER_STRING(FunctionKindName(desc->funckind)));
  SET_NAMES(tmp, names);
  SET_SLOT(ans, Rf_install("kind"), tmp);
  UNPROTECT(2);

  info->ReleaseFuncDesc(desc);
  UNPROTECT(1);
  return(ans);
}

static const char *
FunctionKindName(FUNCKIND kind)
{
  const char *name = "";
  switch(kind) {
  case FUNC_VIRTUAL:
    name = "virtual";
    break;
  case FUNC_PUREVIRTUAL:
    name = "purevirtual";
    break;
  case FUNC_NONVIRTUAL:
    name = "nonvirtual";
    break;
  case FUNC_STATIC:
    name = "static";
    break;
  case FUNC_DISPATCH:
    name = "dispatch";
    break;
  }
  return(name);
}

/*  Currently not accessed via a .Call() in the R code. */
SEXP 
R_GetGuid(SEXP s_info)
{
  ITypeInfo *info = (ITypeInfo *) getRReference(s_info);
  TYPEATTR *attr;
  LPOLESTR ostr;
  SEXP ans = NULL_USER_OBJECT;
  HRESULT hr;

  PROTECT(ans = NEW_CHARACTER(2)) ;
  info->GetTypeAttr(&attr);
  hr = StringFromCLSID(attr->guid, &ostr);
  if(hr == S_OK) {
    SET_STRING_ELT(ans, 0, COPY_TO_USER_STRING(FromBstr(ostr)));
  }
  SET_STRING_ELT(ans, 1, COPY_TO_USER_STRING(FromBstr(attr->lpstrSchema)));
  /*  hr = StringFromCLSID(attr->lcid, &ostr);
  if(hr == S_OK) {
    SET_STRING_ELT(ans, 1, COPY_TO_USER_STRING(FromBstr(ostr)));
  }
  */

  info->ReleaseTypeAttr(attr);
  UNPROTECT(1);
  return(ans);
}


SEXP
R_GetNames(SEXP s_info, SEXP sid)
{
  ITypeInfo *info = (ITypeInfo *) getRReference(s_info);
  unsigned int count = 1, n = 0, i;
  BSTR *els;
  HRESULT hr;
  SEXP ans;

  els = (BSTR*)S_alloc(count, sizeof(BSTR));    
  hr = info->GetNames(INTEGER_DATA(sid)[0], els, count, &n);
  if(hr != ERROR_SUCCESS)
    return(NULL_USER_OBJECT);

  PROTECT(ans = NEW_CHARACTER(n));
  for(i = 0; i < count ; i++) {
    SET_STRING_ELT(ans, i, COPY_TO_USER_STRING(FromBstr(els[i])));      
    SysFreeString(els[i]);
  }
  UNPROTECT(1);
  return(ans);
}

static SEXP
GetNames(ITypeInfo *info, MEMBERID id, UINT count, UINT offset)
{
  BSTR *els;
  UINT n =0, i, ctr;
  SEXP ans = R_NilValue;
  HRESULT hr;


  els = (BSTR*)S_alloc(count, sizeof(BSTR));
  hr = info->GetNames(id, els, count+1, &n);

  if(hr == ERROR_SUCCESS) {
    PROTECT(ans = NEW_CHARACTER(count - offset));
    for(i = offset, ctr = 0; i < count ; i++, ctr++) {
      if(els[i]) {
         SET_STRING_ELT(ans, ctr, COPY_TO_USER_STRING(FromBstr(els[i])));      
	 SysFreeString(els[i]);
      }
    }
    for(i = 0; i < offset; i++)
	 SysFreeString(els[i]);

    UNPROTECT(1);
  }
   
  return(ans);
}

SEXP
R_getInvokeEnum()
{
  SEXP ans, names;
  PROTECT(ans = NEW_INTEGER(4));
  PROTECT(names = NEW_CHARACTER(4));

  INTEGER_DATA(ans)[0] = INVOKE_FUNC;
  SET_STRING_ELT(names, 0, COPY_TO_USER_STRING("func"));

  INTEGER_DATA(ans)[1] = INVOKE_PROPERTYGET;
  SET_STRING_ELT(names, 1, COPY_TO_USER_STRING("propertyget"));

  INTEGER_DATA(ans)[2] = INVOKE_PROPERTYPUT;
  SET_STRING_ELT(names, 2, COPY_TO_USER_STRING("propertyput"));

  INTEGER_DATA(ans)[3] = INVOKE_PROPERTYPUTREF;
  SET_STRING_ELT(names, 3, COPY_TO_USER_STRING("propertyputref"));

  SET_NAMES(ans, names);
  UNPROTECT(2);

  return(ans);
}


SEXP
R_parameterStyle(SHORT flags)
{
  SEXP el;
  PROTECT(el = R_newObject("ParameterStyle"));
  if(flags & PARAMFLAG_FIN)
    LOGICAL(GET_SLOT(el, Rf_install("In")))[0] = flags & PARAMFLAG_FIN;
  if(flags & PARAMFLAG_FOUT)
    LOGICAL(GET_SLOT(el, Rf_install("out")))[0] = flags & PARAMFLAG_FOUT;
  if(flags & PARAMFLAG_FLCID)
    LOGICAL(GET_SLOT(el, Rf_install("lcid")))[0] = flags & PARAMFLAG_FLCID;
  if(flags & PARAMFLAG_FRETVAL)
    LOGICAL(GET_SLOT(el, Rf_install("retval")))[0] = flags & PARAMFLAG_FRETVAL;
  if(flags & PARAMFLAG_FOPT)
    LOGICAL(GET_SLOT(el, Rf_install("optional")))[0] = flags & PARAMFLAG_FOPT;

  UNPROTECT(1);
  return(el);
}

SEXP
R_describeParameter(ELEMDESC *edesc, ITypeInfo *info)
{
  SEXP obj;
  PROTECT(obj = R_newObject("ParameterDescription"));

  SET_SLOT(obj, Rf_install("type"), R_describeType(&edesc->tdesc, info));
  SET_SLOT(obj, Rf_install("style"), R_parameterStyle(edesc->paramdesc.wParamFlags));

    //XXX get the default value if there is one.

  UNPROTECT(1);
  return(obj);
}

SEXP
R_describeType(TYPEDESC *desc, ITypeInfo *info)
{
  TYPEDESC *tmp = desc;
  SEXP obj = R_NilValue;
  int nprotect = 0;

  if(tmp->vt == VT_USERDEFINED) {
    PROTECT(obj = R_newObject("TypeDescriptionRef"));nprotect++;
    NUMERIC_DATA(GET_SLOT(obj, Rf_install("reftype")))[0] = tmp->hreftype;    
  }

  if(desc->vt == VT_PTR) {
    int depth = 1;
    tmp = desc->lptdesc;
    while(tmp->vt == VT_PTR) {
      depth++;
      tmp = tmp->lptdesc;
    }

    //XXX
#if 1
    if(tmp->vt == VT_USERDEFINED) {
      return(R_describeType(tmp, info));
    } else {
#endif
    PROTECT(obj = R_newObject("PointerTypeDescription"));
    nprotect++;
    INTEGER_DATA(GET_SLOT(obj, Rf_install("depth")))[0] = depth;
#if 1
    }
#endif
  } 

  if(tmp->vt == VT_CARRAY) {
    tmp = &desc->lpadesc->tdescElem;
  }

  if(tmp->vt == VT_SAFEARRAY) {
    tmp = desc->lptdesc;
  }

  if(obj == R_NilValue) {
    PROTECT(obj = R_newObject("TypeDescription"));
    nprotect++;
  }
  const char *name = getPrimitiveTypeName(tmp->vt);

  SEXP c;
  PROTECT(c = NEW_CHARACTER(1));
  nprotect++;
  SET_STRING_ELT(c, 0, COPY_TO_USER_STRING(name));
  SET_SLOT(obj, Rf_install("name"), c);
  UNPROTECT(nprotect);

  return(obj);
}

const char *
getPrimitiveTypeName(int which)
{
    switch(which) {
        // VARIANT/VARIANTARG compatible types
    case VT_I2: return "short";
    case VT_I4: return "long";
    case VT_R4: return "float";
    case VT_R8: return "double";
    case VT_CY: return "CY";
    case VT_DATE: return "DATE";
    case VT_BSTR: return "BSTR";
    case VT_DISPATCH: return "IDispatch*";
    case VT_ERROR: return "SCODE";
    case VT_BOOL: return "VARIANT_BOOL";
    case VT_VARIANT: return "VARIANT";
    case VT_UNKNOWN: return "IUnknown*";
    case VT_UI1: return "BYTE";
    case VT_DECIMAL: return "DECIMAL";
    case VT_I1: return "char";
    case VT_UI2: return "USHORT";
    case VT_UI4: return "ULONG";
    case VT_I8: return "__int64";
    case VT_UI8: return "unsigned __int64";
    case VT_INT: return "int";
    case VT_UINT: return "UINT";
    case VT_HRESULT: return "HRESULT";
    case VT_VOID: return "void";
    case VT_LPSTR: return "char*";
    case VT_LPWSTR: return "wchar_t*";
    case VT_RECORD:
      return("Record");
      break;
    case VT_FILETIME:
      return("FileTime");
      break;
    case VT_BLOB:
      return("Blob");
      break;
    case VT_STREAM:
      return("Stream");
      break;
    case VT_STORAGE:
      return("Storage");
      break;
    case VT_STREAMED_OBJECT:
      return("Streamed Object");
      break;
    case VT_STORED_OBJECT:
      return("Stored Object");
      break;
    case VT_BLOB_OBJECT:
      return("Blob Object");
      break;
    case VT_CF:
      return("CF");
      break;
    case VT_CLSID:
      return("ClassID");
      break;
    case VT_BSTR_BLOB:
      return("BStringBlob");
      break;
    case VT_VECTOR:
      return("Vector");
      break;
    case VT_SAFEARRAY:
      return("SafeArray");
      break;
    case VT_CARRAY:
      return("CArray");
      break;
    case VT_ARRAY:
      return("Array");
      break;
    case VT_USERDEFINED:
      return("<User Defined>");
      break;
    default:
      fprintf(stderr, "Unhandled case %d\n", which);fflush(stderr);
      return("?");
    }
}

SEXP
R_newFunctionObject(FUNCDESC *desc)
{
  const char *className = "FunctionDescription";

  if(!desc)
    return(mkString(className));

  /* Is there ever a possibility of getting a desc with a mixed type, i.e. ORed elements of the enum.*/
  switch(desc->invkind) {
    case INVOKE_FUNC:
      className = "FunctionInvokeDescription";
      break;
    case INVOKE_PROPERTYGET:
      className = "PropertyGetDescription";
      break;
    case INVOKE_PROPERTYPUT:
      className = "PropertySetDescription";
      break;
    case INVOKE_PROPERTYPUTREF:
      className = "PropertySetRefDescription";
      break;
  }
  return(R_newObject(className));
}

SEXP
R_newVariableObject()
{
  return(R_newObject("VariableDescription"));
}

SEXP
R_newObject(const char *className)
{
  SEXP klass, ans;
  PROTECT(klass = MAKE_CLASS(className));
  ans = duplicate(NEW(klass));
  UNPROTECT(1);
  return(ans);
}

const char * const
getTypeKindName(TYPEKIND val)
{
  switch(val) {
    case TKIND_ENUM:
      return("enum");
      break;
    case TKIND_RECORD:
      return("record");
      break;
    case TKIND_MODULE:
      return("module");
      break;
    case TKIND_INTERFACE:
      return("interface");
      break;
    case TKIND_DISPATCH:
      return("dispatch");
      break;
    case TKIND_COCLASS:
      return("coclass");
      break;
    case TKIND_ALIAS:
      return("alias");
      break;
    case TKIND_UNION:
      return("union");
      break;
    default:
      break;
  }
  return("");
}


SEXP
R_GetFuncNames(SEXP s_info, SEXP s_mid)
{
  MEMBERID mid;
  ITypeInfo *info;
  FUNCDESC *desc;
  int num;
  HRESULT hr;
  SEXP names;

  info = (ITypeInfo *) getRReference(s_info);
  mid = (MEMBERID) NUMERIC_DATA(s_mid)[0];

  hr = info->GetFuncDesc(mid, &desc);
  if(hr != S_OK)
    COMError(hr);

  num = 1 + desc->cParams;
  info->ReleaseFuncDesc(desc);

  names = GetNames(info, mid, num, 0);
  return(names);
}

SEXP
R_GetIDsOfNames(SEXP s_info, SEXP s_names, SEXP isObject)
{
  ITypeInfo *info = NULL;
  IDispatch *obj = NULL;

  MEMBERID *mids;
  HRESULT hr;
  BSTR *strs;
  unsigned int n, i;
  SEXP ans;

  if(LOGICAL_DATA(isObject)[0]) 
   obj = (IDispatch *) getRReference(s_info);
  else
   info = (ITypeInfo *) getRReference(s_info);

  n = Rf_length(s_names);
  if(n < 1) {
   PROBLEM "R_GetIDsOfNames called with empty name list"
   ERROR;
  }

  mids = (MEMBERID *) S_alloc(n, sizeof(MEMBERID));
  strs = (BSTR *) S_alloc(n, sizeof(BSTR));
  if(!strs || !mids) {
    PROBLEM "Can't allocate space"
      ERROR;
  }

  for(i = 0; i < n ; i++) 
    strs[i] = AsBstr(CHAR(STRING_ELT(s_names, i)));

  /* Now ask for all these identifiers. */
  if(LOGICAL_DATA(isObject)[0]) {
    hr = obj->GetIDsOfNames(IID_NULL, strs, n, LOCALE_USER_DEFAULT, mids);   
  } else {
    hr = info->GetIDsOfNames(strs, n, mids);     
#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
    fprintf(stderr, "Status of GetIDsOfNames %X\n", hr);
#endif
  }

  for(i = 0; i < n ; i++) 
    SysFreeString(strs[i]);

  if(hr != S_OK) {
    const char *tmp;
    switch(hr) {
    case STG_E_INSUFFICIENTMEMORY:
      tmp = "out of memory";
      break;
    case E_OUTOFMEMORY:
      tmp = "out of memory";
      break;
    case E_INVALIDARG:
      tmp = "invalid argument";
      break;
    case DISP_E_UNKNOWNNAME:
      tmp = "unknown name";
      break;
    case DISP_E_UNKNOWNLCID:
      tmp = "unknown lcid";
      break;
    default:
      tmp = "?";
      break;
    }
    fprintf(stderr, "error string %s\n", tmp);
    COMError(hr);
  }


  PROTECT(ans = NEW_NUMERIC(n));
  for(i = 0; i < n ; i++) 
    NUMERIC_DATA(ans)[i] = (double) mids[i];
  UNPROTECT(1);

  return(ans);
}


#if 1
 /* Taken from ErrorUtils.cpp in PyWin32 distribution. */
#include "oaidl.h"

#ifndef _countof
#define _countof(array) (sizeof(array)/sizeof(array[0]))
#endif

void GetScodeString(HRESULT hr, LPTSTR buf, int bufSize)
{
	struct HRESULT_ENTRY
	{
		HRESULT hr;
		LPCTSTR lpszName;
	};
	#define MAKE_HRESULT_ENTRY(hr)    { hr, (#hr) }
	static const HRESULT_ENTRY hrNameTable[] =
	{
		MAKE_HRESULT_ENTRY(S_OK),
		MAKE_HRESULT_ENTRY(S_FALSE),

		MAKE_HRESULT_ENTRY(CACHE_S_FORMATETC_NOTSUPPORTED),
		MAKE_HRESULT_ENTRY(CACHE_S_SAMECACHE),
		MAKE_HRESULT_ENTRY(CACHE_S_SOMECACHES_NOTUPDATED),
		MAKE_HRESULT_ENTRY(CONVERT10_S_NO_PRESENTATION),
		MAKE_HRESULT_ENTRY(DATA_S_SAMEFORMATETC),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_CANCEL),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_DROP),
		MAKE_HRESULT_ENTRY(DRAGDROP_S_USEDEFAULTCURSORS),
		MAKE_HRESULT_ENTRY(INPLACE_S_TRUNCATED),
		MAKE_HRESULT_ENTRY(MK_S_HIM),
		MAKE_HRESULT_ENTRY(MK_S_ME),
		MAKE_HRESULT_ENTRY(MK_S_MONIKERALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(MK_S_REDUCED_TO_SELF),
		MAKE_HRESULT_ENTRY(MK_S_US),
		MAKE_HRESULT_ENTRY(OLE_S_MAC_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(OLE_S_STATIC),
		MAKE_HRESULT_ENTRY(OLE_S_USEREG),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_CANNOT_DOVERB_NOW),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_S_LAST),
		MAKE_HRESULT_ENTRY(STG_S_CONVERTED),
		MAKE_HRESULT_ENTRY(VIEW_S_ALREADY_FROZEN),

		MAKE_HRESULT_ENTRY(E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(E_NOTIMPL),
		MAKE_HRESULT_ENTRY(E_OUTOFMEMORY),
		MAKE_HRESULT_ENTRY(E_INVALIDARG),
		MAKE_HRESULT_ENTRY(E_NOINTERFACE),
		MAKE_HRESULT_ENTRY(E_POINTER),
		MAKE_HRESULT_ENTRY(E_HANDLE),
		MAKE_HRESULT_ENTRY(E_ABORT),
		MAKE_HRESULT_ENTRY(E_FAIL),
		MAKE_HRESULT_ENTRY(E_ACCESSDENIED),

		MAKE_HRESULT_ENTRY(CACHE_E_NOCACHE_UPDATED),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_BAD_DATA),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_CLOSE),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_EMPTY),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_OPEN),
		MAKE_HRESULT_ENTRY(CLIPBRD_E_CANT_SET),
		MAKE_HRESULT_ENTRY(CO_E_ALREADYINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_APPDIDNTREG),
		MAKE_HRESULT_ENTRY(CO_E_APPNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_APPSINGLEUSE),
		MAKE_HRESULT_ENTRY(CO_E_BAD_PATH),
		MAKE_HRESULT_ENTRY(CO_E_CANTDETERMINECLASS),
		MAKE_HRESULT_ENTRY(CO_E_CLASS_CREATE_FAILED),
		MAKE_HRESULT_ENTRY(CO_E_CLASSSTRING),
		MAKE_HRESULT_ENTRY(CO_E_DLLNOTFOUND),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINAPP),
		MAKE_HRESULT_ENTRY(CO_E_ERRORINDLL),
		MAKE_HRESULT_ENTRY(CO_E_IIDSTRING),
		MAKE_HRESULT_ENTRY(CO_E_NOTINITIALIZED),
		MAKE_HRESULT_ENTRY(CO_E_OBJISREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTCONNECTED),
		MAKE_HRESULT_ENTRY(CO_E_OBJNOTREG),
		MAKE_HRESULT_ENTRY(CO_E_OBJSRV_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SCM_ERROR),
		MAKE_HRESULT_ENTRY(CO_E_SCM_RPC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_EXEC_FAILURE),
		MAKE_HRESULT_ENTRY(CO_E_SERVER_STOPPING),
		MAKE_HRESULT_ENTRY(CO_E_WRONGOSFORAPP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_BITMAP_TO_DIB),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_GET),
		MAKE_HRESULT_ENTRY(CONVERT10_E_OLESTREAM_PUT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_DIB_TO_BITMAP),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_FMT),
		MAKE_HRESULT_ENTRY(CONVERT10_E_STG_NO_STD_STREAM),
		MAKE_HRESULT_ENTRY(DISP_E_ARRAYISLOCKED),
		MAKE_HRESULT_ENTRY(DISP_E_BADCALLEE),
		MAKE_HRESULT_ENTRY(DISP_E_BADINDEX),
		MAKE_HRESULT_ENTRY(DISP_E_BADPARAMCOUNT),
		MAKE_HRESULT_ENTRY(DISP_E_BADVARTYPE),
		MAKE_HRESULT_ENTRY(DISP_E_EXCEPTION),
		MAKE_HRESULT_ENTRY(DISP_E_MEMBERNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_NONAMEDARGS),
		MAKE_HRESULT_ENTRY(DISP_E_NOTACOLLECTION),
		MAKE_HRESULT_ENTRY(DISP_E_OVERFLOW),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTFOUND),
		MAKE_HRESULT_ENTRY(DISP_E_PARAMNOTOPTIONAL),
		MAKE_HRESULT_ENTRY(DISP_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNINTERFACE),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(DISP_E_UNKNOWNNAME),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_ALREADYREGISTERED),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(DRAGDROP_E_NOTREGISTERED),
		MAKE_HRESULT_ENTRY(DV_E_CLIPFORMAT),
		MAKE_HRESULT_ENTRY(DV_E_DVASPECT),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE),
		MAKE_HRESULT_ENTRY(DV_E_DVTARGETDEVICE_SIZE),
		MAKE_HRESULT_ENTRY(DV_E_FORMATETC),
		MAKE_HRESULT_ENTRY(DV_E_LINDEX),
		MAKE_HRESULT_ENTRY(DV_E_NOIVIEWOBJECT),
		MAKE_HRESULT_ENTRY(DV_E_STATDATA),
		MAKE_HRESULT_ENTRY(DV_E_STGMEDIUM),
		MAKE_HRESULT_ENTRY(DV_E_TYMED),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTOOLSPACE),
		MAKE_HRESULT_ENTRY(INPLACE_E_NOTUNDOABLE),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_LINK),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_ROOT),
		MAKE_HRESULT_ENTRY(MEM_E_INVALID_SIZE),
		MAKE_HRESULT_ENTRY(MK_E_CANTOPENFILE),
		MAKE_HRESULT_ENTRY(MK_E_CONNECTMANUALLY),
		MAKE_HRESULT_ENTRY(MK_E_ENUMERATION_FAILED),
		MAKE_HRESULT_ENTRY(MK_E_EXCEEDEDDEADLINE),
		MAKE_HRESULT_ENTRY(MK_E_INTERMEDIATEINTERFACENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(MK_E_INVALIDEXTENSION),
		MAKE_HRESULT_ENTRY(MK_E_MUSTBOTHERUSER),
		MAKE_HRESULT_ENTRY(MK_E_NEEDGENERIC),
		MAKE_HRESULT_ENTRY(MK_E_NO_NORMALIZED),
		MAKE_HRESULT_ENTRY(MK_E_NOINVERSE),
		MAKE_HRESULT_ENTRY(MK_E_NOOBJECT),
		MAKE_HRESULT_ENTRY(MK_E_NOPREFIX),
		MAKE_HRESULT_ENTRY(MK_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBINDABLE),
		MAKE_HRESULT_ENTRY(MK_E_NOTBOUND),
		MAKE_HRESULT_ENTRY(MK_E_SYNTAX),
		MAKE_HRESULT_ENTRY(MK_E_UNAVAILABLE),
		MAKE_HRESULT_ENTRY(OLE_E_ADVF),
		MAKE_HRESULT_ENTRY(OLE_E_ADVISENOTSUPPORTED),
		MAKE_HRESULT_ENTRY(OLE_E_BLANK),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_BINDTOSOURCE),
		MAKE_HRESULT_ENTRY(OLE_E_CANT_GETMONIKER),
		MAKE_HRESULT_ENTRY(OLE_E_CANTCONVERT),
		MAKE_HRESULT_ENTRY(OLE_E_CLASSDIFF),
		MAKE_HRESULT_ENTRY(OLE_E_ENUM_NOMORE),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDHWND),
		MAKE_HRESULT_ENTRY(OLE_E_INVALIDRECT),
		MAKE_HRESULT_ENTRY(OLE_E_NOCACHE),
		MAKE_HRESULT_ENTRY(OLE_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(OLE_E_NOSTORAGE),
		MAKE_HRESULT_ENTRY(OLE_E_NOT_INPLACEACTIVE),
		MAKE_HRESULT_ENTRY(OLE_E_NOTRUNNING),
		MAKE_HRESULT_ENTRY(OLE_E_OLEVERB),
		MAKE_HRESULT_ENTRY(OLE_E_PROMPTSAVECANCELLED),
		MAKE_HRESULT_ENTRY(OLE_E_STATIC),
		MAKE_HRESULT_ENTRY(OLE_E_WRONGCOMPOBJ),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_INVALIDVERB),
		MAKE_HRESULT_ENTRY(OLEOBJ_E_NOVERBS),
		MAKE_HRESULT_ENTRY(REGDB_E_CLASSNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_IIDNOTREG),
		MAKE_HRESULT_ENTRY(REGDB_E_INVALIDVALUE),
		MAKE_HRESULT_ENTRY(REGDB_E_KEYMISSING),
		MAKE_HRESULT_ENTRY(REGDB_E_READREGDB),
		MAKE_HRESULT_ENTRY(REGDB_E_WRITEREGDB),
		MAKE_HRESULT_ENTRY(RPC_E_ATTEMPTED_MULTITHREAD),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_CANCELED),
		MAKE_HRESULT_ENTRY(RPC_E_CALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_AGAIN),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INASYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_INEXTERNALCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTCALLOUT_ININPUTSYNCCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTPOST_INSENDCALL),
		MAKE_HRESULT_ENTRY(RPC_E_CANTTRANSMIT_CALL),
		MAKE_HRESULT_ENTRY(RPC_E_CHANGED_MODE),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_CLIENT_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_CONNECTION_TERMINATED),
		MAKE_HRESULT_ENTRY(RPC_E_DISCONNECTED),
		MAKE_HRESULT_ENTRY(RPC_E_FAULT),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_CALLDATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_DATAPACKET),
		MAKE_HRESULT_ENTRY(RPC_E_INVALID_PARAMETER),
		MAKE_HRESULT_ENTRY(RPC_E_INVALIDMETHOD),
		MAKE_HRESULT_ENTRY(RPC_E_NOT_REGISTERED),
		MAKE_HRESULT_ENTRY(RPC_E_OUT_OF_RESOURCES),
		MAKE_HRESULT_ENTRY(RPC_E_RETRY),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_CANTUNMARSHAL_DATA),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVER_DIED_DNE),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_REJECTED),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERCALL_RETRYLATER),
		MAKE_HRESULT_ENTRY(RPC_E_SERVERFAULT),
		MAKE_HRESULT_ENTRY(RPC_E_SYS_CALL_FAILED),
		MAKE_HRESULT_ENTRY(RPC_E_THREAD_NOT_INIT),
		MAKE_HRESULT_ENTRY(RPC_E_UNEXPECTED),
		MAKE_HRESULT_ENTRY(RPC_E_WRONG_THREAD),
		MAKE_HRESULT_ENTRY(STG_E_ABNORMALAPIEXIT),
		MAKE_HRESULT_ENTRY(STG_E_ACCESSDENIED),
		MAKE_HRESULT_ENTRY(STG_E_CANTSAVE),
		MAKE_HRESULT_ENTRY(STG_E_DISKISWRITEPROTECTED),
		MAKE_HRESULT_ENTRY(STG_E_EXTANTMARSHALLINGS),
		MAKE_HRESULT_ENTRY(STG_E_FILEALREADYEXISTS),
		MAKE_HRESULT_ENTRY(STG_E_FILENOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_INSUFFICIENTMEMORY),
		MAKE_HRESULT_ENTRY(STG_E_INUSE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFLAG),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHANDLE),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDHEADER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDNAME),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPARAMETER),
		MAKE_HRESULT_ENTRY(STG_E_INVALIDPOINTER),
		MAKE_HRESULT_ENTRY(STG_E_LOCKVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_MEDIUMFULL),
		MAKE_HRESULT_ENTRY(STG_E_NOMOREFILES),
		MAKE_HRESULT_ENTRY(STG_E_NOTCURRENT),
		MAKE_HRESULT_ENTRY(STG_E_NOTFILEBASEDSTORAGE),
		MAKE_HRESULT_ENTRY(STG_E_OLDDLL),
		MAKE_HRESULT_ENTRY(STG_E_OLDFORMAT),
		MAKE_HRESULT_ENTRY(STG_E_PATHNOTFOUND),
		MAKE_HRESULT_ENTRY(STG_E_READFAULT),
		MAKE_HRESULT_ENTRY(STG_E_REVERTED),
		MAKE_HRESULT_ENTRY(STG_E_SEEKERROR),
		MAKE_HRESULT_ENTRY(STG_E_SHAREREQUIRED),
		MAKE_HRESULT_ENTRY(STG_E_SHAREVIOLATION),
		MAKE_HRESULT_ENTRY(STG_E_TOOMANYOPENFILES),
		MAKE_HRESULT_ENTRY(STG_E_UNIMPLEMENTEDFUNCTION),
		MAKE_HRESULT_ENTRY(STG_E_UNKNOWN),
		MAKE_HRESULT_ENTRY(STG_E_WRITEFAULT),
		MAKE_HRESULT_ENTRY(TYPE_E_AMBIGUOUSNAME),
		MAKE_HRESULT_ENTRY(TYPE_E_BADMODULEKIND),
		MAKE_HRESULT_ENTRY(TYPE_E_BUFFERTOOSMALL),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTCREATETMPFILE),
		MAKE_HRESULT_ENTRY(TYPE_E_CANTLOADLIBRARY),
		MAKE_HRESULT_ENTRY(TYPE_E_CIRCULARTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_DLLFUNCTIONNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_DUPLICATEID),
		MAKE_HRESULT_ENTRY(TYPE_E_ELEMENTNOTFOUND),
		MAKE_HRESULT_ENTRY(TYPE_E_INCONSISTENTPROPFUNCS),
		MAKE_HRESULT_ENTRY(TYPE_E_INVALIDSTATE),
		MAKE_HRESULT_ENTRY(TYPE_E_INVDATAREAD),
		MAKE_HRESULT_ENTRY(TYPE_E_IOERROR),
		MAKE_HRESULT_ENTRY(TYPE_E_LIBNOTREGISTERED),
		MAKE_HRESULT_ENTRY(TYPE_E_NAMECONFLICT),
		MAKE_HRESULT_ENTRY(TYPE_E_OUTOFBOUNDS),
		MAKE_HRESULT_ENTRY(TYPE_E_QUALIFIEDNAMEDISALLOWED),
		MAKE_HRESULT_ENTRY(TYPE_E_REGISTRYACCESS),
		MAKE_HRESULT_ENTRY(TYPE_E_SIZETOOBIG),
		MAKE_HRESULT_ENTRY(TYPE_E_TYPEMISMATCH),
		MAKE_HRESULT_ENTRY(TYPE_E_UNDEFINEDTYPE),
		MAKE_HRESULT_ENTRY(TYPE_E_UNKNOWNLCID),
		MAKE_HRESULT_ENTRY(TYPE_E_UNSUPFORMAT),
		MAKE_HRESULT_ENTRY(TYPE_E_WRONGTYPEKIND),
		MAKE_HRESULT_ENTRY(VIEW_E_DRAW),

#if NOT_AVAILABLE
		MAKE_HRESULT_ENTRY(CONNECT_E_NOCONNECTION),
		MAKE_HRESULT_ENTRY(CONNECT_E_ADVISELIMIT),
		MAKE_HRESULT_ENTRY(CONNECT_E_CANNOTCONNECT),
		MAKE_HRESULT_ENTRY(CONNECT_E_OVERRIDDEN),
#endif

#ifndef NO_PYCOM_IPROVIDECLASSINFO
		MAKE_HRESULT_ENTRY(CLASS_E_NOAGGREGATION),
		MAKE_HRESULT_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
#endif // NO_PYCOM_IPROVIDECLASSINFO

#ifndef MS_WINCE // ??
#if AVAILABLE
		MAKE_HRESULT_ENTRY(CTL_E_ILLEGALFUNCTIONCALL      ),
		MAKE_HRESULT_ENTRY(CTL_E_OVERFLOW                 ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFMEMORY              ),
		MAKE_HRESULT_ENTRY(CTL_E_DIVISIONBYZERO           ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTRINGSPACE         ),
		MAKE_HRESULT_ENTRY(CTL_E_OUTOFSTACKSPACE          ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAMEORNUMBER      ),
		MAKE_HRESULT_ENTRY(CTL_E_FILENOTFOUND             ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILEMODE              ),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYOPEN          ),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEIOERROR            ),
		MAKE_HRESULT_ENTRY(CTL_E_FILEALREADYEXISTS        ),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDLENGTH          ),
		MAKE_HRESULT_ENTRY(CTL_E_DISKFULL                 ),
		MAKE_HRESULT_ENTRY(CTL_E_BADRECORDNUMBER          ),
		MAKE_HRESULT_ENTRY(CTL_E_BADFILENAME              ),
		MAKE_HRESULT_ENTRY(CTL_E_TOOMANYFILES             ),
		MAKE_HRESULT_ENTRY(CTL_E_DEVICEUNAVAILABLE        ),
		MAKE_HRESULT_ENTRY(CTL_E_PERMISSIONDENIED         ),
		MAKE_HRESULT_ENTRY(CTL_E_DISKNOTREADY             ),
		MAKE_HRESULT_ENTRY(CTL_E_PATHFILEACCESSERROR      ),
		MAKE_HRESULT_ENTRY(CTL_E_PATHNOTFOUND             ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPATTERNSTRING     ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDUSEOFNULL         ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDFILEFORMAT        ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYVALUE     ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPROPERTYARRAYINDEX),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTEDATRUNTIME ),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTSUPPORTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_NEEDPROPERTYARRAYINDEX   ),
		MAKE_HRESULT_ENTRY(CTL_E_SETNOTPERMITTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTEDATRUNTIME ),
		MAKE_HRESULT_ENTRY(CTL_E_GETNOTSUPPORTED          ),
		MAKE_HRESULT_ENTRY(CTL_E_PROPERTYNOTFOUND         ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDCLIPBOARDFORMAT   ),
		MAKE_HRESULT_ENTRY(CTL_E_INVALIDPICTURE           ),
		MAKE_HRESULT_ENTRY(CTL_E_PRINTERERROR             ),
		MAKE_HRESULT_ENTRY(CTL_E_CANTSAVEFILETOTEMP       ),
		MAKE_HRESULT_ENTRY(CTL_E_SEARCHTEXTNOTFOUND       ),
		MAKE_HRESULT_ENTRY(CTL_E_REPLACEMENTSTOOLONG      ),
#endif
#endif // MS_WINCE
	};
	#undef MAKE_HRESULT_ENTRY

	// first ask the OS to give it to us..
	// ### should we get the Unicode version instead?
	int numCopied = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, buf, bufSize, NULL );
	if (numCopied>0) {
		if (numCopied<bufSize) {
			// trim trailing crap
			if (numCopied>2 && (buf[numCopied-2]=='\n'||buf[numCopied-2]=='\r'))
				buf[numCopied-2] = '\0';
		}
		return;
	}

	// else look for it in the table
	for (unsigned int i = 0; i < _countof(hrNameTable); i++)
	{
		if (hr == hrNameTable[i].hr) {
			strncpy(buf, hrNameTable[i].lpszName, bufSize);
			return;
		}
	}
	// not found - make one up
	sprintf(buf, ("OLE error 0x%08x"), (unsigned int) hr);
}
#endif



#if 1
/* Copied temporarily from connect.cpp in RDCOMClient. */ 
int
R_getCLSIDFromString(SEXP className, CLSID *classId)
{
  HRESULT hr;
  const char *ptr;
  int status = FALSE;

  BSTR str;
  ptr = CHAR(STRING_ELT(className, 0)); 
  str = AsBstr(ptr);
   
  hr = CLSIDFromString(str, classId);
  if(SUCCEEDED(hr)) {
    SysFreeString(str);
    return(TRUE);
  }

  status = SUCCEEDED(CLSIDFromProgID(str, classId));
  SysFreeString(str);

  return status;
}
#endif

#if 0

/*
 This is not used. See R_QueryPathOfRegTypeLib.
*/ 
SEXP
R_getLibraryLocation(SEXP guid, SEXP version, SEXP language)
{
  SEXP ans;
//XXX   str = SysAllocStringLen(wstr, size);
  BSTR str =  AsBstr("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
  CLSID classId;
  HRESULT hr;

  if(!R_getCLSIDFromString(guid, &classId)) {
    PROBLEM "Can't get class identifier from string"
      ERROR;
  }
  hr = QueryPathOfRegTypeLib(classId, INTEGER(version)[0], INTEGER(version)[1], 0, &str);
 
  ans = mkString(FromBstr(str));
  SysFreeString(str);
  return(ans);
}
#endif


