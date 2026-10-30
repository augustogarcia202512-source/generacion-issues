import mammoth from 'mammoth';

export async function convertDocxToHtml(arrayBuffer) {
  const result = await mammoth.convertToHtml({ arrayBuffer });
  return result.value;
}
