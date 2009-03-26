;;;; -*- Mode: LISP; Syntax: COMMON-LISP; Coding: shift_jis; -*-
(in-package :cl-win32ole-sys)

(cffi:defctype WORD :unsigned-short)
(cffi:defctype DWORD :unsigned-long)
(cffi:defctype SHORT :short)
(cffi:defctype USHORT :unsigned-short)
(cffi:defctype UINT :unsigned-int)
(cffi:defctype LONGLONG :int64)
(cffi:defctype LONG :long)
(cffi:defctype ULONG :unsigned-long)

(cffi:defctype PVOID :pointer "void*")
(cffi:defctype SCODE :long)
(cffi:defctype HRESULT :long)

(cffi:defctype BSTR :pointer)

(cffi:defctype DATE :double)

(cffi:defcstruct GUID
  (Data1 :unsigned-long)
  (Data2 :unsigned-short)
  (Data3 :unsigned-short)
  (Data4 :unsigned-char :count 8))

(cffi:defctype CLSID GUID)

(cffi:defctype LCID DWORD)

(cffi:defctype DISPID :long)

(cffi:defctype MEMBERID DISPID)

(cffi:defcstruct DISPPARAMS
  (rgvarg :pointer)
  (rgdispidNamedArgs :pointer)
  (cArgs UINT)
  (cNamedArgs UINT))

(defvar *null-disp-params*
  (let ((v (cffi:foreign-alloc 'DISPPARAMS)))
    (setf (cffi:foreign-slot-value v 'DISPPARAMS 'rgvarg)
          (cffi-sys:null-pointer)
          (cffi:foreign-slot-value v 'DISPPARAMS 'rgdispidNamedArgs)
          (cffi-sys:null-pointer)
          (cffi:foreign-slot-value v 'DISPPARAMS 'cArgs)
          0
          (cffi:foreign-slot-value v 'DISPPARAMS 'cNamedArgs)
          0)
    v))

(cffi:defctype LPOLESTR :pointer)

(cffi:defctype VARTYPE :unsigned-short)

(cffi:defctype HREFTYPE DWORD)

(cffi:defcstruct IUnknownVtbl
  (QueryInterface :pointer)
  (AddRef :pointer)
  (Release :pointer))

(cffi:defcstruct IUnknown
  (vtbl :pointer))

(cffi:defcstruct IDispatchVtbl
  (QueryInterface :pointer)
  (AddRef :pointer)
  (Release :pointer)
  (GetTypeInfoCount :pointer)
  (GetTypeInfo :pointer)
  (GetIDsOfNames :pointer)
  (Invoke :pointer))

(cffi:defcstruct IDispatch
  (vtbl :pointer))

(cffi:defcstruct ITypeInfoVtbl
  (QueryInterface :pointer)
  (AddRef :pointer)
  (Release :pointer)
  (GetTypeAttr :pointer)
  (GetTypeComp :pointer)
  (GetFuncDesc :pointer)
  (GetVarDesc :pointer)
  (GetNames :pointer)
  (GetRefTypeOfImplType :pointer)
  (GetImplTypeFlags :pointer)
  (GetIDsOfNames :pointer)
  (Invoke :pointer)
  (GetDocumentation :pointer)
  (GetDllEntry :pointer)
  (GetRefTypeInfo :pointer)
  (AddressOfMember :pointer)
  (CreateInstance :pointer)
  (GetMops :pointer)
  (GetContainingTypeLib :pointer)
  (ReleaseTypeAttr :pointer)
  (ReleaseFuncDesc :pointer)
  (ReleaseVarDesc :pointer))

(cffi:defcstruct ITypeInfo
  (vtbl :pointer))

(cffi:defcstruct ITypeLibVtbl
  (QueryInterface :pointer)
  (AddRef :pointer)
  (Release :pointer)
  (GetTypeInfoCount :pointer)
  (GetTypeInfo :pointer)
  (GetTypeInfoType :pointer)
  (GetTypeInfoOfGuid :pointer)
  (GetLibAttr :pointer)
  (GetTypeComp :pointer)
  (GetDocumentation :pointer)
  (IsName :pointer)
  (FindName :pointer)
  (ReleaseTLibAttr :pointer))

(cffi:defcstruct ITypeLib
  (vtbl :pointer))

(cffi:defcstruct TYPEDESC
  (desc :pointer)
  (vt VARTYPE))

(cffi:defcstruct IDLDESC
  (dwReserved :pointer)
  (wIDLFlags :unsigned-short))

(cffi:defcstruct TYPEATTR
  (guid GUID)
  (lcid LCID)
  (dwReserved DWORD)
  (memidConstructor MEMBERID)
  (memidDestructor MEMBERID)
  (lpstrSchema LPOLESTR)
  (cbSizeInstance ULONG)
  (typekind :int)
  (cFuncs WORD)
  (cVars WORD)
  (cImplTypes WORD)
  (cbSizeVft WORD)
  (cbAlignment WORD)
  (wTypeFlags WORD)
  (wMajorVerNum WORD)
  (wMinorVerNum WORD)
  (tdescAlias TYPEDESC)
  (idldescType IDLDESC))



(cffi:defcstruct PARAMDESC
  (pparamdescex :pointer)
  (wParamFlags USHORT))


(cffi:defcstruct ELEMDESC
  (tdesc TYPEDESC)
  (idl-param-desc PARAMDESC))           ; union { IDLDESC idldesc;
                                        ;         PARAMDESC paramdesc; }

(cffi:defcstruct FUNCDESC
  (memid MEMBERID)
  (lprgscode :pointer)
  (lprgelemdescParam :pointer)
  (funckind :int)
  (invkind :int)
  (callconv :int)
  (cParams SHORT)
  (cParamsOpt SHORT)
  (oVft SHORT)
  (cScodes SHORT)
  (elemdescFunc ELEMDESC)
  (wFuncFlags WORD))


(cffi:defcstruct VARDESC
  (memid MEMBERID)
  (lpstrSchema LPOLESTR)
  (value :pointer)                      ; union { ULONG oInst;
                                        ;         VARIAN* lpvarValue };
  (elemdescVar ELEMDESC)
  (wVarFlags WORD)
  (varkind :int))

(cffi:defcstruct SAFEARRAYBOUND
  (elements ULONG)
  (l-bound  LONG))

(cffi:defcstruct EXCEPINFO
  (wCode  WORD)
  (wReserved  WORD)
  (bstrSource  BSTR)
  (bstrDescription  BSTR)
  (bstrHelpFile  BSTR)
  (dwHelpContext DWORD)
  (pvReserved PVOID)
  (pfnDeferredFillIn PVOID) ;HRESULT (__stdcall *pfnDeferredFillIn)(struct tagEXCEPINFO *)
  (scode SCODE))
