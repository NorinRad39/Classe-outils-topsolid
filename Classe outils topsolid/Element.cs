using System.Collections.Generic;
using System.Linq;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace OutilsTs
{
    /// <summary>
    /// Classe repr�sentant un �l�ment TopSolid avec ses propri�t�s �tendues.
    /// Encapsule un ElementId et fournit un acc�s simplifi� aux propri�t�s de l'�l�ment,
    /// y compris la d�tection automatique des shapes et le calcul de leur volume.
    /// </summary>
    /// <remarks>
    /// Namespace: OutilsTs  
    /// Assembly: OutilsTs (in OutilsTs.dll)
    /// </remarks>
    /// <example>
    /// <code>
    /// // Cr�er un �l�ment � partir d'un ElementId
    /// ElementId elementId = TSH.Elements.SearchByName(docId, "MonElement");
    /// Element element = new Element(elementId);
    /// 
    /// // V�rifier si c'est un shape et obtenir son volume
    /// if (element.IsShape)
    /// {
    ///     Console.WriteLine($"Volume: {element.VolumeMm3} mm�");
    /// }
    /// </code>
    /// </example>
    public class Element
    {
        #region Champs priv�s
        private ElementId elementId;
        private string friendlyName;
        private string typeFullName;
        private bool isShape;
        private double? volume; // Nullable car tous les �l�ments ne sont pas des shapes
        #endregion

        #region Constructeurs
        /// <summary>
        /// Initialise un nouvel �l�ment � partir de son ElementId.
        /// R�cup�re automatiquement toutes les propri�t�s de l'�l�ment lors de l'initialisation.
        /// </summary>
        /// <param name="elementId">Identifiant de l'�l�ment TopSolid.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Cr�er un �l�ment
        /// ElementId id = TSH.Elements.SearchByName(docId, "Electrode");
        /// Element element = new Element(id);
        /// Console.WriteLine($"Nom: {element.FriendlyName}");
        /// </code>
        /// </example>
        public Element(ElementId elementId)
        {
            // Assigne l'identifiant de l'�l�ment
            this.elementId = elementId;
            
            // Initialise automatiquement toutes les propri�t�s de l'�l�ment
            // (nom, type, volume si c'est un shape)
            InitializeElement();
        }
        #endregion

        #region Propri�t�s publiques
        /// <summary>
        /// Obtient l'identifiant de l'�l�ment TopSolid.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// ElementId id = element.ElementId;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="ElementId"/>
        /// Identifiant de l'�l�ment.
        /// </returns>
        public ElementId ElementId 
        { 
            get => elementId; 
        }

        /// <summary>
        /// Obtient le nom convivial de l'�l�ment.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Le nom convivial est le nom affich� dans l'interface TopSolid.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// string nom = element.FriendlyName; // Ex: "Electrode_1"
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom convivial de l'�l�ment, ou "Nom inconnu" en cas d'erreur.
        /// </returns>
        public string FriendlyName 
        { 
            get => friendlyName; 
        }

        /// <summary>
        /// Obtient le nom complet du type de l'�l�ment.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Le nom complet du type inclut le namespace complet de la classe TopSolid.
        /// Exemple: "TopSolid.Kernel.DB.D3.Shapes.Prism"
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// string type = element.TypeFullName;
        /// // Ex: "TopSolid.Kernel.DB.D3.Shapes.Prism"
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom complet du type de l'�l�ment, ou "Type inconnu" en cas d'erreur.
        /// </returns>
        public string TypeFullName 
        { 
            get => typeFullName; 
        }

        /// <summary>
        /// Indique si l'�l�ment est un shape (forme 3D).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Un shape est identifi� par son type qui commence par "TopSolid.Kernel.DB.D3.Shapes.".
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.IsShape)
        /// {
        ///     Console.WriteLine("C'est un shape !");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si l'�l�ment est un shape, sinon <c>false</c>.
        /// </returns>
        public bool IsShape 
        { 
            get => isShape; 
        }

        /// <summary>
        /// Obtient le volume du shape en m�tres cubes (m�).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Cette propri�t� est <c>null</c> si l'�l�ment n'est pas un shape.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.Volume.HasValue)
        /// {
        ///     Console.WriteLine($"Volume: {element.Volume.Value} m�");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="T:System.Nullable1"/>
        /// Volume en m�, ou <c>null</c> si l'�l�ment n'est pas un shape.
        /// </returns>
        public double? Volume 
        { 
            get => volume; 
        }

        /// <summary>
        /// Obtient le volume du shape en millim�tres cubes (mm�).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Cette propri�t� effectue automatiquement la conversion de m� en mm�.
        /// Conversion: 1 m� = 1 000 000 000 mm�.
        /// Cette propri�t� est <c>null</c> si l'�l�ment n'est pas un shape.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.VolumeMm3.HasValue)
        /// {
        ///     Console.WriteLine($"Volume: {element.VolumeMm3.Value:F2} mm�");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="T:System.Nullable1"/>
        /// Volume en mm�, ou <c>null</c> si l'�l�ment n'est pas un shape.
        /// </returns>
        public double? VolumeMm3 
        { 
            // Cast explicite n�cessaire pour C# 7.3
            get => volume.HasValue ? (double?)(volume.Value * 1_000_000_000) : null; 
        }
        #endregion

        #region M�thodes priv�es
        /// <summary>
        /// Initialise les propri�t�s de l'�l�ment en interrogeant l'API TopSolid.
        /// Cette m�thode est appel�e automatiquement par le constructeur.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// �tapes d'initialisation :
        /// 1. R�cup�ration du nom convivial (FriendlyName)
        /// 2. R�cup�ration du type complet (TypeFullName)
        /// 3. D�tection si c'est un shape (analyse du nom du type)
        /// 4. Si c'est un shape, r�cup�ration du volume
        /// 
        /// En cas d'erreur sur une propri�t�, une valeur par d�faut est assign�e.
        /// </remarks>
        private void InitializeElement()
        {
            // V�rifier si l'ElementId est valide
            // Si vide, on arr�te l'initialisation
            if (elementId.IsEmpty) return;

            // --- �tape 1 : R�cup�ration du nom convivial ---
            try
            {
                // R�cup�rer le nom convivial depuis l'API TopSolid
                friendlyName = TSH.Elements.GetFriendlyName(elementId);
            }
            catch
            {
                // En cas d'erreur, on assigne un nom par d�faut
                friendlyName = "Nom inconnu";
            }

            // --- �tape 2 : R�cup�ration du type complet ---
            try
            {
                // R�cup�rer le nom complet du type depuis l'API TopSolid
                // Ex: "TopSolid.Kernel.DB.D3.Shapes.Prism"
                typeFullName = TSH.Elements.GetTypeFullName(elementId);
            }
            catch
            {
                // En cas d'erreur, on assigne un type par d�faut
                typeFullName = "Type inconnu";
            }

            // --- �tape 3 & 4 : D�tection shape et r�cup�ration du volume ---
            try
            {
                // V�rifier si c'est un shape en analysant le nom du type
                // Un shape a un type qui commence par "TopSolid.Kernel.DB.D3.Shapes."
                isShape = !string.IsNullOrEmpty(typeFullName) && 
                          typeFullName.StartsWith("TopSolid.Kernel.DB.D3.Shapes.");

                // Si c'est un shape, r�cup�rer son volume en m�
                if (isShape)
                {
                    volume = TSH.Shapes.GetShapeVolume(elementId);
                }
                else
                {
                    // Si ce n'est pas un shape, le volume est null
                    volume = null;
                }
            }
            catch
            {
                // En cas d'erreur, on consid�re que ce n'est pas un shape
                isShape = false;
                volume = null;
            }
        }
        #endregion
    }

    /// <summary>
    /// M�thodes d'extension pour faciliter la manipulation des �l�ments TopSolid.
    /// Fournit des m�thodes utilitaires pour convertir et trier les �l�ments.
    /// </summary>
    /// <remarks>
    /// Namespace: OutilsTs  
    /// Assembly: OutilsTs (in OutilsTs.dll)
    /// </remarks>
    /// <example>
    /// <code>
    /// // Convertir une liste d'ElementId en �l�ments enrichis
    /// List&lt;ElementId&gt; ids = TSH.Shapes.GetShapes(docId);
    /// List&lt;Element&gt; elements = ids.ToElements();
    /// 
    /// // Trier les shapes par volume
    /// List&lt;Element&gt; shapesTries = elements.GetShapesSortedByVolume();
    /// </code>
    /// </example>
    public static class ElementExtensions
    {
        /// <summary>
        /// Convertit une liste d'ElementId en liste d'objets Element enrichis.
        /// Chaque ElementId est encapsul� dans un objet Element qui fournit
        /// un acc�s simplifi� aux propri�t�s de l'�l�ment.
        /// </summary>
        /// <param name="elementIds">Liste des identifiants d'�l�ments � convertir.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// Cette m�thode d'extension permet une conversion fluide des ElementId en objets Element.
        /// Si la liste source est null, retourne une liste vide.
        /// </remarks>
        /// <example>
        /// <code>
        /// // R�cup�rer tous les shapes d'un document
        /// List&lt;ElementId&gt; shapeIds = TSH.Shapes.GetShapes(docId);
        /// 
        /// // Convertir en objets Element
        /// List&lt;Element&gt; elements = shapeIds.ToElements();
        /// 
        /// // Utiliser les propri�t�s enrichies
        /// foreach (var element in elements)
        /// {
        ///     Console.WriteLine($"{element.FriendlyName}: {element.VolumeMm3} mm�");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{Element}"/>
        /// Liste des objets Element, ou liste vide si elementIds est null.
        /// </returns>
        public static List<Element> ToElements(this List<ElementId> elementIds)
        {
            // Cr�er une nouvelle liste pour stocker les �l�ments convertis
            List<Element> elements = new List<Element>();
            
            // Si la liste source est null, retourner une liste vide
            if (elementIds == null) return elements;
            
            // Convertir chaque ElementId en objet Element
            foreach (var id in elementIds)
            {
                elements.Add(new Element(id));
            }
            
            return elements;
        }

        /// <summary>
        /// Filtre uniquement les shapes d'une liste d'�l�ments et les trie par volume.
        /// Par d�faut, le tri est d�croissant (du plus gros au plus petit volume).
        /// </summary>
        /// <param name="elements">Liste des �l�ments � filtrer et trier.</param>
        /// <param name="descending">
        /// <c>true</c> pour un tri d�croissant (par d�faut), 
        /// <c>false</c> pour un tri croissant.
        /// </param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// Cette m�thode effectue deux op�rations :
        /// 1. Filtre les �l�ments pour ne garder que les shapes ayant un volume
        /// 2. Trie les shapes par volume (d�croissant ou croissant)
        /// 
        /// Les �l�ments qui ne sont pas des shapes ou qui n'ont pas de volume sont exclus.
        /// </remarks>
        /// <example>
        /// <code>
        /// // R�cup�rer et convertir les �l�ments
        /// List&lt;ElementId&gt; ids = TSH.Shapes.GetShapes(docId);
        /// List&lt;Element&gt; elements = ids.ToElements();
        /// 
        /// // Trier par volume d�croissant (plus gros en premier)
        /// List&lt;Element&gt; shapesTries = elements.GetShapesSortedByVolume();
        /// 
        /// // Afficher avec num�rotation
        /// for (int i = 0; i &lt; shapesTries.Count; i++)
        /// {
        ///     Console.WriteLine($"{i + 1}. {shapesTries[i].FriendlyName} - {shapesTries[i].VolumeMm3:F2} mm�");
        /// }
        /// 
        /// // Trier par volume croissant (plus petit en premier)
        /// List&lt;Element&gt; shapesCroissant = elements.GetShapesSortedByVolume(descending: false);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{Element}"/>
        /// Liste des shapes tri�s par volume.
        /// </returns>
        public static List<Element> GetShapesSortedByVolume(this List<Element> elements, bool descending = true)
        {
            // Filtrer pour ne garder que les shapes ayant un volume
            // LINQ Where : filtre les �l�ments selon la condition
            // ToList : convertit le r�sultat en List<Element>
            var shapes = elements.Where(e => e.IsShape && e.Volume.HasValue).ToList();
            
            // Trier par volume
            if (descending)
            {
                // Tri d�croissant : b compar� � a (inversion)
                // Le plus gros volume sera en premier
                shapes.Sort((a, b) => b.Volume.Value.CompareTo(a.Volume.Value));
            }
            else
            {
                // Tri croissant : a compar� � b (ordre normal)
                // Le plus petit volume sera en premier
                shapes.Sort((a, b) => a.Volume.Value.CompareTo(b.Volume.Value));
            }
            
            return shapes;
        }
    }
}