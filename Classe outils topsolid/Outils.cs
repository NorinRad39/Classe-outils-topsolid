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

namespace OutilsTs
{
    public class Document
    {
        private DocumentId docId;
        private string docNomTxt;
        private string docExtention;
        private PdmObjectId docPdmObject;
        private List<ElementId> docParameters = new List<ElementId>();
        private List<ElementId> docOperations = new List<ElementId>();
        private ElementId docCommentaireSysId;
        private ElementId docDescriptionSysId;
        private bool docDerived;
        private bool docIsElectrode;
        private static Dictionary<DocumentId, string> _cacheNomDocuments = new Dictionary<DocumentId, string>();

        public Document()
        {
            // Récupérer le document courant
            try
            {
                DocId = TSH.Documents.EditedDocument;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la récupération du document courant : {ex.Message}");
                // Optionnellement, initialisez avec une valeur par défaut ou null
            }
        }

        public Document(DocumentId docId)
        {
            DocId = docId;
        }

        public DocumentId DocId
        {
            get => docId;
            set
            {
                docId = value;
                try
                {
                    docNomTxt = TSH.Documents.GetName(docId) ?? "Nom inconnu";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération du nom du document : {ex.Message}");
                    docNomTxt = "Erreur";
                }
                try
                {
                    docPdmObject = TSH.Documents.GetPdmObject(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération de l'objet PDM : {ex.Message}");
                    docPdmObject = PdmObjectId.Empty; // Valeur par défaut en cas d'erreur
                }
                try
                {
                    docParameters = TSH.Parameters.GetParameters(docId) ?? new List<ElementId>();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération des paramètres : {ex.Message}");
                    docParameters = new List<ElementId>(); // Liste vide en cas d'erreur
                }
                try
                {
                    docOperations = TSH.Operations.GetOperations(docId) ?? new List<ElementId>();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération des opérations : {ex.Message}");
                    docOperations = new List<ElementId>(); // Liste vide en cas d'erreur
                }
                try
                {
                    docCommentaireSysId = TSH.Parameters.GetCommentParameter(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération du commentaire système : {ex.Message}");
                    docCommentaireSysId = ElementId.Empty; // Valeur par défaut en cas d'erreur
                }
                try
                {
                    docDescriptionSysId = TSH.Parameters.GetDescriptionParameter(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération de la description système : {ex.Message}");
                    docDescriptionSysId = ElementId.Empty; // Valeur par défaut en cas d'erreur
                }
                try
                {
                    docDerived = TSHD.Tools.IsDerived(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la vérification du caractère dérivé du document : {ex.Message}");
                    docDerived = false; // Valeur par défaut en cas d'erreur
                }
                try
                {
                    docIsElectrode = IsElectrode(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la vérification si le document est une électrode : {ex.Message}");
                    docIsElectrode = false; // Valeur par défaut en cas d'erreur
                }
                try
                {
                    TSH.Pdm.GetType(docPdmObject, out docExtention);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération de l'extension du document : {ex.Message}");
                    docExtention = null; // Valeur par défaut en cas d'erreur
                }
            }
        }

        public string DocNomTxt { get => docNomTxt; private set => docNomTxt = value; }
        public string DocExtention { get => docExtention; private set => docExtention = value; }
        public PdmObjectId DocPdmObject
        {
            get => docPdmObject;
            set
            {
                docPdmObject = value;
                TSH.Pdm.GetType(docPdmObject, out docExtention);
            }
        }
        public List<ElementId> DocParameters { get => docParameters; private set => docParameters = value; }
        public List<ElementId> DocOperations { get => docOperations; private set => docOperations = value; }
        public ElementId DocCommentaireSysId { get => docCommentaireSysId; private set => docCommentaireSysId = value; }
        public ElementId DocDescriptionSysId { get => docDescriptionSysId; private set => docDescriptionSysId = value; }
        public bool DocDerived { get => docDerived; private set => docDerived = value; }
        public bool DocIsElectrode { get => docIsElectrode; private set => docIsElectrode = value; }
        public List<ElementId> CamOperations => GetCamOperations(DocId);
        public bool DocuCam => IsDocuCam(DocId);
        public string OP => NumOP(DocId);

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

        private bool IsElectrode(DocumentId document)
        {
            try
            {
                ElementId systemParametersFolder = TSH.Parameters.GetSystemParametersFolder(document);
                List<ElementId> parameterListe = TSH.Elements.GetConstituents(systemParametersFolder);

                if (parameterListe == null || parameterListe.Count == 0)
                {
                    return false;
                }

                foreach (ElementId param in parameterListe)
                {
                    string nom = TSH.Elements.GetName(param);

                    if (nom == "$TopSolid.Kernel.TX.Properties.ElectrodeTemplateAssociation" ||
                        nom == "$TopSolid.Cad.Electrode.DB.Electrodes.ElectrodeDocument")
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                string errorMsg = $"Erreur lors de la vérification de l'électrode : {ex.Message}";
                Console.WriteLine(errorMsg);
                return false;
            }
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
