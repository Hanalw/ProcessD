export const fetchSiteColumns = async (siteurl: string): Promise<any> => {
  const filter = "Group ne '_Hidden' and Hidden eq false and (TypeAsString eq 'TaxonomyFieldTypeMulti' or TypeAsString eq 'TaxonomyFieldType')";
  const endpoint = `${siteurl}/_api/web/fields?$filter=${filter}`;
  const response = await fetch(endpoint, {
    method: 'GET',
    headers: {
      'Accept': 'application/json;odata=verbose'
    }
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch site columns: ${response.statusText}`);
  }
  const data = await response.json();
  return data.d.results;
};