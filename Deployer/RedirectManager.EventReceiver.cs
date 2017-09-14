using System;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using Microsoft.SharePoint;

namespace Features.RedirectManager
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("a1f3adf4-2da6-4081-80a4-eb59007c73df")]
    public class FeatureReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSite targetSite = properties.Feature.Parent as SPSite;
            LocalizeSiteLookupFields(targetSite.ID, new string[] { "PageReference" });
        }

        public void LocalizeSiteLookupFields(Guid targetSiteId, string[] fields)
        {
            using (SPSite site = new SPSite(targetSiteId.ToString()))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    foreach (var field in fields)
                    {
                        LocalizeSiteLookupField(web, field);
                    }
                }
            }

        }
        public void LocalizeSiteLookupField(SPWeb web, string fieldName)
        {
            SPField lookupField = web.Fields.TryGetFieldByStaticName(fieldName);

            if (lookupField != null)
            {
                String transformedFieldXmlConfig = CreateLocalizedLookupFieldXmlConfig(lookupField.SchemaXml, web);
                lookupField.SchemaXml = transformedFieldXmlConfig;
            }
        }

        /**
         * 
         * Sample Lookup Field xml config:
         *<Field
         *  Group="Redirect Manager Columns"
         *  ID="{e49e6631-f9a3-4984-ad9a-dab46ec52015}"
         *  Type="LookupMulti"
         *  Name="PageReference"
         *  DisplayName="Page Reference"
         *  StaticName="PageReference"
         *  Required="FALSE"
         *  EnforceUniqueValues="FALSE"
         *  List="Pages"
         *  ShowField="Title"
         *  Mult="TRUE"
         *  Sortable="FALSE"
         *  UnlimitedLengthInDocumentLibrary="FALSE"
         *  Version="1">
         *</Field>
         *  
         */
        public string CreateLocalizedLookupFieldXmlConfig(string fieldXmlConfig, SPWeb web)
        {
            string transformedFieldXmlConfig = fieldXmlConfig;
            // Getting field xml configuration
            XDocument fieldXmlConfigDocument = XDocument.Parse(fieldXmlConfig);

            // Get the root element of the field definition
            XElement root = fieldXmlConfigDocument.Root;

            string fieldType = root.Attribute("Type")?.Value.ToLower();
            string fieldName = root.Attribute("Name")?.Value;
            XAttribute targetLookupListUrl = root.Attribute("List");
            XAttribute sourceIDAttribute = root.Attribute("SourceID");

            if (fieldType != "lookup" || fieldType != "lookupmulti")
            {
                throw new Exception("The specified lookup field name '{fieldName}' has type '{fieldType}'.  Expecting field type value of 'Lookup' or 'LookupMulti'.");
            }

            // Check if target lookup list definition exits
            if (targetLookupListUrl != null)
            {
                // Use the configured list url attribute value to find actual target list GUID 
                string targetListUrl = targetLookupListUrl.Value;

                // Get the target lookup list folder to acquire it's parent list id
                SPFolder listFolder = web.GetFolder(targetListUrl);
                if (listFolder != null)
                {
                    // Replace the url with the id of the list folder parent
                    targetLookupListUrl.Value = listFolder.ParentListId.ToString();

                    // Setting the souce id of the schema
                    if (sourceIDAttribute != null)
                    {
                        // Replace the sourceid with the correct webid
                        sourceIDAttribute.Value = web.ID.ToString();
                    }

                    transformedFieldXmlConfig = fieldXmlConfigDocument.ToString();
                }
            }
            return transformedFieldXmlConfig;
        }

    }

    // Uncomment the method below to handle the event raised before a feature is deactivated.

    //public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
    //{
    //}


    // Uncomment the method below to handle the event raised after a feature has been installed.

    //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
    //{
    //}


    // Uncomment the method below to handle the event raised before a feature is uninstalled.

    //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
    //{
    //}

    // Uncomment the method below to handle the event raised when a feature is upgrading.

    //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
    //{
    //}
}
