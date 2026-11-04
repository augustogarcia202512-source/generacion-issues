import {createOlePackage as buildOlePackage} from './excel-ole-core.js';
import iconBytes from './ole-icon.emf';
import templateBytes from './ole-package-template.bin';

export function createOlePackage(imageBytes, extension) {
  return buildOlePackage(imageBytes, extension, templateBytes);
}

export {iconBytes};
