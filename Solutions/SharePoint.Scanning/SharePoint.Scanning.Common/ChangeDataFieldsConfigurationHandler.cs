using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SharePoint.Scanning.Common
{
    public class FieldElement : ConfigurationElement
    {
        [ConfigurationProperty("fieldName", IsRequired = true)]
        public string FieldName
        {
            get { return (string)base["fieldName"]; }
            set { base["fieldName"] = value; }
        }

        [ConfigurationProperty("profilePropertyName", IsRequired = true)]
        public string ProfilePropertyName
        {
            get { return (string)base["profilePropertyName"]; }
            set { base["profilePropertyName"] = value; }
        }

        [ConfigurationProperty("taxonomyTermSetName", IsRequired = true)]
        public string TaxonomyTermSetName
        {
            get { return (string)base["taxonomyTermSetName"]; }
            set { base["taxonomyTermSetName"] = value; }
        }

        internal string Key
        {
            get { return string.Format("{0}|{1}|{2}", FieldName, ProfilePropertyName, TaxonomyTermSetName); }
        }
    }

    [ConfigurationCollection(typeof(FieldElement), AddItemName = "field", CollectionType = ConfigurationElementCollectionType.BasicMap)]
    public class FieldElementCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new FieldElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((FieldElement)element).Key;
        }

        public void Add( FieldElement element)
        {
            BaseAdd(element);
        }

        public void Clear()
        {
            BaseClear();
        }

        public List<string> GetFieldNameValues()
        {
            var list = new List<string>();
            for(int index = 0; index < this.Count; index++)
            {
                var fieldInfo = (FieldElement)BaseGet(index);
                list.Add(fieldInfo.FieldName);
            }

            return list;
        }

        public List<string> GetProfilePropertyValues()
        {
            var list = new List<string>();
            for(int index =0; index < this.Count; index++)
            {
                var fieldInfo = (FieldElement)BaseGet(index);
                list.Add(fieldInfo.ProfilePropertyName);
            }
            return list;
        }

        public List<string> GetTaxonomyTermSetNames()
        {
            var list = new List<string>();
            for (int index = 0; index < this.Count; index++)
            {
                var fieldInfo = (FieldElement)BaseGet(index);
                list.Add(fieldInfo.TaxonomyTermSetName);
            }
            return list;
        }

        public int IndexOf (FieldElement element)
        {
            return BaseIndexOf(element);
        }
        public void Remove(FieldElement element)
        {
            if (BaseIndexOf(element) >= 0)
            {
                BaseRemove(element.Key);
            }
        }

        public void RemoveAt(int index)
        {
            BaseRemoveAt(index);
        }

        public FieldElement this[int index]
        {
            get { return (FieldElement)BaseGet(index); }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                }
                BaseAdd(index, value);
            }
        }
    }

    public class FieldDataSection : ConfigurationSection
    {
        private static readonly ConfigurationProperty _propFieldInfo = new ConfigurationProperty(
            null, 
            typeof(FieldElementCollection), 
            null, 
            ConfigurationPropertyOptions.IsDefaultCollection);

        private static readonly ConfigurationProperty _propTaxonomyTermGroup = new ConfigurationProperty(
            "TermGroup", typeof(string), "People", ConfigurationPropertyOptions.IsRequired);

      
        private static ConfigurationPropertyCollection _properties = new ConfigurationPropertyCollection();

        static FieldDataSection()
        {
            _properties.Add(_propTaxonomyTermGroup);
            _properties.Add(_propFieldInfo);
        }

        [ConfigurationProperty("TermGroup")]
        public String TermGroup
        {
            get { return (String)base[_propTaxonomyTermGroup]; }
        }

        [ConfigurationProperty("", Options = ConfigurationPropertyOptions.IsDefaultCollection)]
        public FieldElementCollection Fields
        {
            get { return (FieldElementCollection)base[_propFieldInfo]; }
        }
    }
}
