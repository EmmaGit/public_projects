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
    
    public partial class Report
    {
        public int Id { get; set; }
        public System.DateTime DateTime { get; set; }
        public Nullable<System.TimeSpan> Duration { get; set; }
        public Nullable<int> TotalChecked { get; set; }
        public Nullable<int> TotalErrors { get; set; }
        public string ResultPath { get; set; }
        public string Module { get; set; }
        public Nullable<int> ScheduleId { get; set; }
        public string Author { get; set; }
    
        public virtual Email Email { get; set; }
        public virtual Schedule Schedule { get; set; }
    }
}
