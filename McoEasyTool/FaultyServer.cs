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
    
    public partial class FaultyServer : Server
    {
        public FaultyServer()
        {
            this.FaultyServer_Reports = new HashSet<FaultyServer_Report>();
        }
    
        public string IdSite { get; set; }
        public string Site { get; set; }
    
        public virtual ICollection<FaultyServer_Report> FaultyServer_Reports { get; set; }
    }
}
