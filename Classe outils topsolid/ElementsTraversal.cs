using System;
using System.Collections.Generic;
using TopSolid.Kernel.Automating;

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour la traversée récursive d'éléments TopSolid
    /// </summary>
    /// <example>
    /// Exemples d'utilisation :
    /// 
    /// // Pour GetConstituents (récupérer tous les constituants récursivement)
    /// var constituents = ElementsTraversal.GetAllDescendants(
    ///     paralleOperationsActive, 
    ///     elem => TopSolidHost.Elements.GetConstituents(elem)
    /// );
    /// 
    /// // Pour GetChildren (récupérer tous les enfants récursivement)
    /// var children = ElementsTraversal.GetAllDescendants(
    ///     operations,
    ///     elem => TopSolidHost.Operations.GetChildren(elem),
    ///     maxDepth: 3 // Limiter à 3 niveaux
    /// );
    /// 
    /// // Avec filtrage (exemple : récupérer uniquement les solides)
    /// var solids = ElementsTraversal.GetAllDescendantsFiltered(
    ///     rootElement,
    ///     elem => TopSolidHost.Elements.GetConstituents(elem),
    ///     elem => TopSolidHost.Elements.GetTypeFullName(elem).Contains("Solid")
    /// );
    /// </example>
    public static class ElementsTraversal
    {
        /// <summary>
        /// Récupère récursivement tous les descendants d'un élément en utilisant une fonction de navigation personnalisée
        /// </summary>
        /// <param name="rootElement">L'élément racine à partir duquel commencer la recherche</param>
        /// <param name="getChildrenFunc">Fonction qui retourne les enfants d'un élément (ex: GetChildren, GetConstituents)</param>
        /// <param name="maxDepth">Profondeur maximale de récursion (null = illimitée)</param>
        /// <returns>Liste de tous les descendants trouvés</returns>
        public static List<ElementId> GetAllDescendants(
            ElementId rootElement, 
            Func<ElementId, List<ElementId>> getChildrenFunc,
            int? maxDepth = null)
        {
            if (rootElement.IsEmpty || getChildrenFunc == null)
            {
                return new List<ElementId>();
            }

            var results = new List<ElementId>();
            var visited = new HashSet<ElementId>(); // Protection contre les cycles

            GetAllDescendantsRecursive(rootElement, getChildrenFunc, results, visited, 0, maxDepth);

            return results;
        }

        /// <summary>
        /// Récupère récursivement tous les descendants d'une liste d'éléments
        /// </summary>
        /// <param name="rootElements">Liste d'éléments racines</param>
        /// <param name="getChildrenFunc">Fonction qui retourne les enfants d'un élément</param>
        /// <param name="maxDepth">Profondeur maximale de récursion (null = illimitée)</param>
        /// <returns>Liste de tous les descendants trouvés</returns>
        public static List<ElementId> GetAllDescendants(
            List<ElementId> rootElements,
            Func<ElementId, List<ElementId>> getChildrenFunc,
            int? maxDepth = null)
        {
            if (rootElements == null || rootElements.Count == 0 || getChildrenFunc == null)
            {
                return new List<ElementId>();
            }

            var results = new List<ElementId>();
            var visited = new HashSet<ElementId>();

            foreach (var element in rootElements)
            {
                if (!element.IsEmpty)
                {
                    GetAllDescendantsRecursive(element, getChildrenFunc, results, visited, 0, maxDepth);
                }
            }

            return results;
        }

        /// <summary>
        /// Méthode récursive interne pour parcourir l'arborescence
        /// </summary>
        private static void GetAllDescendantsRecursive(
            ElementId currentElement,
            Func<ElementId, List<ElementId>> getChildrenFunc,
            List<ElementId> results,
            HashSet<ElementId> visited,
            int currentDepth,
            int? maxDepth)
        {
            // Vérifier si on a déjà visité cet élément (évite les cycles)
            if (visited.Contains(currentElement))
            {
                return;
            }

            // Marquer comme visité
            visited.Add(currentElement);

            // Vérifier la profondeur maximale
            if (maxDepth.HasValue && currentDepth >= maxDepth.Value)
            {
                return;
            }

            // Récupérer les enfants via la fonction fournie
            List<ElementId> children = null;
            try
            {
                children = getChildrenFunc(currentElement);
            }
            catch
            {
                // Si l'appel échoue, on continue sans enfants
                return;
            }

            // Traiter chaque enfant
            if (children != null && children.Count > 0)
            {
                foreach (var child in children)
                {
                    if (!child.IsEmpty)
                    {
                        // Ajouter l'enfant aux résultats
                        results.Add(child);

                        // Appel récursif pour les descendants de cet enfant
                        GetAllDescendantsRecursive(child, getChildrenFunc, results, visited, currentDepth + 1, maxDepth);
                    }
                }
            }
        }

        /// <summary>
        /// Récupère récursivement tous les descendants filtrés par un prédicat
        /// </summary>
        /// <param name="rootElement">L'élément racine</param>
        /// <param name="getChildrenFunc">Fonction qui retourne les enfants</param>
        /// <param name="filterFunc">Fonction de filtrage (retourne true pour inclure l'élément)</param>
        /// <param name="maxDepth">Profondeur maximale</param>
        /// <returns>Liste des descendants filtrés</returns>
        public static List<ElementId> GetAllDescendantsFiltered(
            ElementId rootElement,
            Func<ElementId, List<ElementId>> getChildrenFunc,
            Func<ElementId, bool> filterFunc,
            int? maxDepth = null)
        {
            var allDescendants = GetAllDescendants(rootElement, getChildrenFunc, maxDepth);

            if (filterFunc == null)
            {
                return allDescendants;
            }

            var filtered = new List<ElementId>();
            foreach (var element in allDescendants)
            {
                try
                {
                    if (filterFunc(element))
                    {
                        filtered.Add(element);
                    }
                }
                catch
                {
                    // Ignorer les éléments qui causent des erreurs lors du filtrage
                }
            }

            return filtered;
        }
    }
}