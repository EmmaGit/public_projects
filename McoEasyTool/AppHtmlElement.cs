//------------------------------------------------------------------------------
// <auto-generated>
//     Ce code a été généré à partir d'un modèle.
//
//     Des modifications manuelles apportées à ce fichier peuvent conduire à un comportement inattendu de votre application.
//     Les modifications manuelles apportées à ce fichier sont remplacées si le code est régénéré.
// </auto-generated>
//------------------------------------------------------------------------------

namespace McoEasyTool
{
    using System;
    using System.Collections.Generic;
    
    public partial class AppHtmlElement
    {
        public int Id { get; set; }
        public string TagName { get; set; }
        public string AttrId { get; set; }
        public string AttrName { get; set; }
        public string AttrClass { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public int ApplicationId { get; set; }
        public string AttrXpath { get; set; }
    
        public virtual Application Application { get; set; }
    }
}
