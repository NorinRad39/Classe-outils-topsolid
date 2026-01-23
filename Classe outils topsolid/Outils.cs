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

/*
================================================================================
Liste des méthodes et propriétés disponibles dans la classe Document (OutilsTs)
================================================================================

Constructeurs :
- Document()
    // Récupère le document courant TopSolid.
    // Syntaxe :
    // var doc = new Document();

- Document(DocumentId docId)
    // Initialise un document à partir de son identifiant.
    // Syntaxe :
    // var doc = new Document(monDocId);

Propriétés principales :
- string DocNomTxt
    // Nom du document.
    // Syntaxe :
    // string nom = doc.DocNomTxt;

- string DocExtention
    // Extension/type PDM du document.
    // Syntaxe :
    // string ext = doc.DocExtention;

- PdmObjectId DocPdmObject
    // Objet PDM associé au document.
    // Syntaxe :
    // PdmObjectId pdm = doc.DocPdmObject;

- List<ElementId> DocParameters
    // Liste des paramètres du document.
    // Syntaxe :
    // var parametres = doc.DocParameters;

- List<ElementId> DocOperations
    // Liste des opérations du document.
    // Syntaxe :
    // var operations = doc.DocOperations;

- ElementId DocCommentaireSysId
    // Identifiant du paramètre commentaire système.
    // Syntaxe :
    // var commentaireId = doc.DocCommentaireSysId;

- ElementId DocDescriptionSysId
    // Identifiant du paramètre description système.
    // Syntaxe :
    // var descriptionId = doc.DocDescriptionSysId;

- bool DocDerived
    // Indique si le document est dérivé.
    // Syntaxe :
    // bool estDerive = doc.DocDerived;

- bool DocIsElectrode
    // Indique si le document est une électrode.
    // Syntaxe :
    // bool estElectrode = doc.DocIsElectrode;

- List<ElementId> CamOperations
    // Liste des opérations CAM (si document CAM).
    // Syntaxe :
    // var camOps = doc.CamOperations;

- bool DocuCam
    // Indique si le document est un document CAM.
    // Syntaxe :
    // bool estCam = doc.DocuCam;

- string OP
    // Numéro d’OP (si présent).
    // Syntaxe :
    // string op = doc.OP;

================================================================================
Pour toute nouvelle méthode ou propriété, ajoutez-la à cette liste pour faciliter la découverte.
================================================================================
*/

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour la gestion et l'accès aux propriétés, paramètres, opérations et métadonnées
    /// des documents TopSolid.
    /// </summary>
    public class Document
    {
        #region Champs privés
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
        #endregion

        #region Constructeurs
        /// <summary>
        /// Initialise une nouvelle instance de la classe <see cref="Document"/> représentant le document courant TopSolid.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Récupérer le document courant
        /// var doc = new Document();
        /// </code>
        /// </example>
        public Document()
        {
            try
            {
                docId = TSH.Documents.EditedDocument;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la récupération du document courant : {ex.Message}");
            }
        }

        /// <summary>
        /// Initialise une nouvelle instance de la classe <see cref="Document"/> à partir de l'identifiant d'un document TopSolid.
        /// </summary>
        /// <param name="docId">Identifiant du document TopSolid.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Récupérer un document par son identifiant
        /// var doc = new Document(monDocId);
        /// </code>
        /// </example>
        public Document(DocumentId docId)
        {
            DocId = docId;  // ← Ceci assigne à une propriété DocId
        }
        #endregion

        #region Propriétés publiques
        /// <summary>
        /// Obtient ou définit l'identifiant du document TopSolid.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// DocumentId id = doc.DocId;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="DocumentId"/>
        /// Identifiant du document.
        /// </returns>
        public DocumentId DocId 
        { 
            get => docId; 
            set => docId = value; 
        }

        /// <summary>
        /// Obtient le nom du document TopSolid.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// string nom = doc.DocNomTxt;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom du document.
        /// </returns>
        public string DocNomTxt { get => docNomTxt; private set => docNomTxt = value; }

        /// <summary>
        /// Obtient l’extension/type PDM du document.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// string ext = doc.DocExtention;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Extension ou type PDM du document.
        /// </returns>
        public string DocExtention { get => docExtention; private set => docExtention = value; }

        /// <summary>
        /// Obtient l’objet PDM associé au document.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// PdmObjectId pdm = doc.DocPdmObject;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="PdmObjectId"/>
        /// Objet PDM du document.
        /// </returns>
        public PdmObjectId DocPdmObject
        {
            get => docPdmObject;
            set
            {
                docPdmObject = value;
                TSH.Pdm.GetType(docPdmObject, out docExtention);
            }
        }

        /// <summary>
        /// Obtient la liste des paramètres du document.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// var parametres = doc.DocParameters;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des paramètres du document.
        /// </returns>
        public List<ElementId> DocParameters { get => docParameters; private set => docParameters = value; }

        /// <summary>
        /// Obtient la liste des opérations du document.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// var operations = doc.DocOperations;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des opérations du document.
        /// </returns>
        public List<ElementId> DocOperations { get => docOperations; private set => docOperations = value; }

        /// <summary>
        /// Obtient l’identifiant du paramètre commentaire système.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// var commentaireId = doc.DocCommentaireSysId;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="ElementId"/>
        /// Identifiant du paramètre commentaire système.
        /// </returns>
        public ElementId DocCommentaireSysId { get => docCommentaireSysId; private set => docCommentaireSysId = value; }

        /// <summary>
        /// Obtient l’identifiant du paramètre description système.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// var descriptionId = doc.DocDescriptionSysId;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="ElementId"/>
        /// Identifiant du paramètre description système.
        /// </returns>
        public ElementId DocDescriptionSysId { get => docDescriptionSysId; private set => docDescriptionSysId = value; }

        /// <summary>
        /// Indique si le document est dérivé.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// bool estDerive = doc.DocDerived;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si le document est dérivé, sinon <c>false</c>.
        /// </returns>
        public bool DocDerived { get => docDerived; private set => docDerived = value; }

        /// <summary>
        /// Indique si le document est une électrode.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// bool estElectrode = doc.DocIsElectrode;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si le document est une électrode, sinon <c>false</c>.
        /// </returns>
        public bool DocIsElectrode { get => docIsElectrode; private set => docIsElectrode = value; }

        /// <summary>
        /// Obtient la liste des opérations CAM du document (si document CAM).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// var camOps = doc.CamOperations;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des opérations CAM.
        /// </returns>
        public List<ElementId> CamOperations => GetCamOperations(DocId);

        /// <summary>
        /// Indique si le document est un document CAM.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// bool estCam = doc.DocuCam;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si le document est un document CAM, sinon <c>false</c>.
        /// </returns>
        public bool DocuCam => IsDocuCam(DocId);

        /// <summary>
        /// Obtient le numéro d’OP du document (si présent).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var doc = new Document();
        /// string op = doc.OP;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Numéro d’OP du document.
        /// </returns>
        public string OP => NumOP(DocId);
        #endregion

        #region Méthodes privées
        /// <summary>
        /// Vérifie si l'identifiant de document est valide.
        /// </summary>
        /// <param name="document">Identifiant du document à vérifier.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// bool isValid = IsValidDocument(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si l'identifiant est valide, sinon <c>false</c>.
        /// </returns>
        private bool IsValidDocument(DocumentId document)
        {
            if (document == DocumentId.Empty)
            {
                UiHelper.ShowMessage("Erreur : DocumentId vide.", "Erreur", MessageType.Error);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Gère l'affichage et la journalisation des exceptions.
        /// </summary>
        /// <param name="ex">Exception à traiter.</param>
        /// <param name="message">Message d'erreur à afficher et journaliser.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// try
        /// {
        ///     // ...
        /// }
        /// catch (Exception ex)
        /// {
        ///     HandleException(ex, "Erreur personnalisée");
        /// }
        /// </code>
        /// </example>
        private void HandleException(Exception ex, string message)
        {
            LogError(message, ex);
            UiHelper.ShowMessage($"{message}\nDétails : {ex.Message}", "Erreur", MessageType.Error);
        }

        /// <summary>
        /// Ajoute une entrée dans le fichier errors.log.
        /// </summary>
        /// <param name="message">Message d'erreur à journaliser.</param>
        /// <param name="ex">Exception associée.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// LogError("Erreur lors de l'opération", ex);
        /// </code>
        /// </example>
        private void LogError(string message, Exception ex)
        {
            string logPath = "errors.log";
            File.AppendAllText(logPath, $"{DateTime.Now}: {message} - {ex.Message}\n");
        }

        /// <summary>
        /// Détermine si le document est une électrode.
        /// </summary>
        /// <param name="document">Identifiant du document à tester.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// bool isElectrode = IsElectrode(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si le document est une électrode, sinon <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Récupère les opérations du document.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var ops = GetOperations(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des opérations du document.
        /// </returns>
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

        /// <summary>
        /// Récupère les opérations CAM du document.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var camOps = GetCamOperations(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des opérations CAM.
        /// </returns>
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

        /// <summary>
        /// Récupère les paramètres du document.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var parameters = GetDocuParameters(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{ElementId}"/>
        /// Liste des paramètres du document.
        /// </returns>
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

        /// <summary>
        /// Récupère le nom du document (avec cache).
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// string nom = GetNomDocu(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom du document.
        /// </returns>
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

        /// <summary>
        /// Récupère l'extension/type du document.
        /// </summary>
        /// <param name="pdmObject">Objet PDM du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// string ext = GetExtension(pdmObject);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Extension ou type du document.
        /// </returns>
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

        /// <summary>
        /// Indique si le document est un document CAM.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// bool isCam = IsDocuCam(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si le document est un document CAM, sinon <c>false</c>.
        /// </returns>
        private bool IsDocuCam(DocumentId document)
        {
            if (!IsValidDocument(document)) return false;
            return TSCH.Documents.IsCam(document);
        }

        /// <summary>
        /// Récupère l'objet PDM du document.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// PdmObjectId pdm = PdmObjectDocu(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="PdmObjectId"/>
        /// Objet PDM du document.
        /// </returns>
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

        /// <summary>
        /// Récupère le numéro d’OP du document.
        /// </summary>
        /// <param name="document">Identifiant du document.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// string op = NumOP(docId);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Numéro d’OP du document.
        /// </returns>
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
        #endregion
    }
}
