using Classe_outils_topsolid;
using System;
using System.Collections.Generic;
using System.IO;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TopSolid.Kernel.Automating;
using TSCH = TopSolid.Cam.NC.Kernel.Automating.TopSolidCamHost;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using TopSolid.Cad.Drafting.Automating;

namespace TSCamPgmRename_WPF
{
    public class Document
    {
        public DocumentId DocId { get; set; }

        public List<ElementId> Operations => GetOperations(DocId);
        public List<ElementId> CamOperations => GetCamOperations(DocId);
        public List<ElementId> Parametres => GetDocuParameters(DocId);
        public string Nom => GetNomDocu(DocId);
        public string Extension => GetExtension(PdmObject);
        public bool DocuCam => IsDocuCam(DocId);
        public PdmObjectId PdmObject => PdmObjectDocu(DocId);
        public string OP => NumOP(DocId);

        private static Dictionary<DocumentId, string> _cacheNomDocuments = new Dictionary<DocumentId, string>();

        public Document() { }

        public Document(DocumentId docId)
        {
            DocId = docId;
        }

        private bool IsValidDocument(DocumentId document)
        {
            if (document == DocumentId.Empty)
            {
                UiHelper.ShowMessage("Erreur : DocumentId vide.", "Erreur", MessageType.Error);
                return false;
            }
            return true;
        }

        private void HandleException(Exception ex, string message)
        {
            LogError(message, ex);
            UiHelper.ShowMessage($"{message}\nDétails : {ex.Message}", "Erreur", MessageType.Error);
        }

        private void LogError(string message, Exception ex)
        {
            string logPath = "errors.log";
            File.AppendAllText(logPath, $"{DateTime.Now}: {message} - {ex.Message}\n");
        }

        private List<ElementId> GetOperations(DocumentId document)
        {
            if (!IsValidDocument(document)) return new List<ElementId>();

            try
            {
                return TSH.Operations.GetOperations(document);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer les opérations.");
                return new List<ElementId>();
            }
        }

        private List<ElementId> GetCamOperations(DocumentId document)
        {
            if (!IsValidDocument(document) || !TSCH.Documents.IsCam(document))
                return new List<ElementId>();

            try
            {
                return TSCH.Operations.GetOperations(document);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer les opérations CAM.");
                return new List<ElementId>();
            }
        }

        private List<ElementId> GetDocuParameters(DocumentId document)
        {
            if (!IsValidDocument(document)) return new List<ElementId>();

            try
            {
                return TSH.Parameters.GetParameters(document);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer les paramètres.");
                return new List<ElementId>();
            }
        }

        private string GetNomDocu(DocumentId document)
        {
            if (!IsValidDocument(document)) return string.Empty;

            if (_cacheNomDocuments.ContainsKey(document))
                return _cacheNomDocuments[document];

            try
            {
                string nom = TSH.Documents.GetName(document);
                _cacheNomDocuments[document] = nom;
                return nom;
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer le nom du document.");
                return string.Empty;
            }
        }

        private string GetExtension(PdmObjectId pdmObject)
        {
            try
            {
                TSH.Pdm.GetType(pdmObject, out string extension);
                return extension;
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer l'extension du document.");
                return string.Empty;
            }
        }

        private bool IsDocuCam(DocumentId document)
        {
            if (!IsValidDocument(document)) return false;
            return TSCH.Documents.IsCam(document);
        }

        private PdmObjectId PdmObjectDocu(DocumentId document)
        {
            if (!IsValidDocument(document)) return new PdmObjectId();

            try
            {
                return TSH.Documents.GetPdmObject(document);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer l'ID PDM du document.");
                return new PdmObjectId();
            }
        }

        private string NumOP(DocumentId document)
        {
            if (!IsValidDocument(document)) return string.Empty;

            try
            {
                ElementId OP = TSH.Elements.SearchByName(document, "OP");
                return OP != ElementId.Empty ? TSH.Parameters.GetTextValue(OP) : string.Empty;
            }
            catch (Exception ex)
            {
                HandleException(ex, "Impossible de récupérer le numéro OP.");
                return string.Empty;
            }
        }
    }
}
