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
    
    public partial class ReftechServers
    {
        public int IdServeur { get; set; }
        public string NomMachineServeur { get; set; }
        public string NomLogiqueServeur { get; set; }
        public string ServiceMajeur { get; set; }
        public string Environnement { get; set; }
        public string Perimetre { get; set; }
        public string CodeDomaine { get; set; }
        public string IsDedie { get; set; }
        public string IsHauteDispo { get; set; }
        public string IsMaquette { get; set; }
        public string NumSerie { get; set; }
        public string EtatServeur { get; set; }
        public string IdSite { get; set; }
        public string LocalisationPhysique { get; set; }
        public Nullable<int> IdUserExploitLocal { get; set; }
        public string RemarquesServeur { get; set; }
        public System.DateTime DateInsertServeur { get; set; }
        public string IP { get; set; }
        public Nullable<System.DateTime> DateUpdateServeur { get; set; }
        public string IsPCI { get; set; }
        public Nullable<int> IdTypeSupervision { get; set; }
        public Nullable<int> NiveauSupervision { get; set; }
        public Nullable<int> IdExploitantServeur { get; set; }
        public string IdKM { get; set; }
        public Nullable<int> IdServeurParent { get; set; }
        public Nullable<int> IdTypeServeur { get; set; }
    }
}
