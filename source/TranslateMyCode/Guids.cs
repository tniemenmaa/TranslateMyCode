// Guids.cs
// MUST match guids.h
using System;

namespace TomiNiemenmaa.TranslateMyCode
{
    static class GuidList
    {
        public const string guidTranslateMyCodePkgString = "a2391485-c37e-4e43-994d-293ec91857e2";
        public const string guidTranslateMyCodeCmdSetString = "89caf741-9995-4e99-abe1-1de02118d9da";

        public static readonly Guid guidTranslateMyCodeCmdSet = new Guid(guidTranslateMyCodeCmdSetString);
    };
}