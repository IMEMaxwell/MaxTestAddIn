using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Windows.Forms;

namespace MaxTestAddIn
{
    // Macro Feature class
    public class MacroFeature : ISwComFeature
    {
        public dynamic Edit(object app, object modelDoc, object feature)
        {
            // get macro feature data value
            MacroFeatureData swMacroFeatData = default(MacroFeatureData);
            object paramName;
            object paramTypes;
            object paramValues;
            string[] sparamValues;
            string facePID;

            swMacroFeatData = (MacroFeatureData)((Feature)feature).GetDefinition();
            swMacroFeatData.GetParameters(out paramName, out paramTypes, out paramValues);
            sparamValues = (string[])paramValues;
            facePID = sparamValues[0];

            // get solidworks app
            ISldWorks swApp = (ISldWorks)app;

            // get this addin with GUID
            SwAddin swAddin = swApp.GetAddInObject("{175fc051-47ae-446a-9b43-3af101538b9f}");

            // PMP check the group 2 and set focus to it
            MFPMP mfPage = swAddin.mfpage;
            mfPage.Show();
            mfPage.group2.Checked = true;
            mfPage.selection1.SetSelectionFocus();

            // select face
            IModelDoc2 swModel = (IModelDoc2)modelDoc; ;
            ISelectionMgr selMgr = swModel.ISelectionManager;
            IModelDocExtension swModelExt = swModel.Extension; ;
            object swSelObj = GetObjectFromString(swApp, swModel, facePID);
            if (swSelObj != null)
            {
                // cast object to entity
                IEntity faceEntity = (IEntity)swSelObj;

                // check if the casting is valid
                if (faceEntity != null)
                {
                    // select face
                    faceEntity.Select2(false, 0);
                }
                else
                {
                    MessageBox.Show("The object is not an entity");
                }
            }
            else
            {
                MessageBox.Show("Failed to get the object by persist reference");
            }

            return true;
        }

        public dynamic Regenerate(object app, object modelDoc, object feature)
        {
            // default code
            object functionReturnValue = null;
            MacroFeatureData swMacroFeatData = default(MacroFeatureData);
            swMacroFeatData = (MacroFeatureData)((Feature)feature).GetDefinition();
            
            return functionReturnValue;
        }

        public dynamic Security(object app, object modelDoc, object feature)
        {
            return null;
        }

        // function to get object from PID string
        // copy from https://help.solidworks.com/2023/English/api/sldworksapi/Use_Persistent_Reference_Example_CSharp.htm
        public object GetObjectFromString(ISldWorks swApp, IModelDoc2 swModel, string IDstring)
        {
            object functionReturnValue = null;
            ModelDocExtension swModExt = default(ModelDocExtension);
            byte[] ByteStream = new byte[IDstring.Length / 3];
            object vPIDarr = null;
            int nRetval = 0;
            int i = 0;

            swModExt = swModel.Extension;
            for (i = 0; i <= IDstring.Length - 3; i += 3)
            {
                int j;
                j = i / 3;
                ByteStream[j] = Convert.ToByte(IDstring.Substring(i, 3));
            }

            vPIDarr = ByteStream;

            functionReturnValue = swModExt.GetObjectByPersistReference3((vPIDarr), out nRetval);

            return functionReturnValue;

        }
    }
}
