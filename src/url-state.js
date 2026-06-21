export function createViewerUrl(pageUrl, presentationUrl) {
  const result = new URL(pageUrl);
  const params = new URLSearchParams(result.search);
  params.delete('url');

  const existingQuery = params.toString();
  if (!presentationUrl) {
    result.search = existingQuery;
    return result;
  }

  const encodedPresentationUrl = encodeURIComponent(presentationUrl)
    .replaceAll('%3A', ':')
    .replaceAll('%2F', '/');
  result.search = `${existingQuery ? `${existingQuery}&` : ''}url=${encodedPresentationUrl}`;
  return result;
}
