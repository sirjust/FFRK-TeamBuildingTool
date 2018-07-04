using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FFRK_Element_Info.Models
{
    public class Character
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Iteration { get; set; }
        public string Type { get; set; }
        public string ElementType { get; set; }
        public string EnElement { get; set; }
        public string SBType { get; set; }
        public string Level99 { get; set; }
        public string Dived { get; set; }
        public string ExtraInfo { get; set; }
    }
}