import { emuToPixels } from '../../helpers.js';
import { handleParagraphNode } from './paragraphNodeImporter.js';
import { defaultNodeListHandler } from './docxImporter.js';

export const handlePictNode = (params) => {
  const { nodes } = params;

  if (!nodes.length || nodes[0].name !== 'w:p') {
    return { nodes: [], consumed: 0 };
  }

  const [pNode] = nodes;
  const run = pNode.elements?.find((el) => el.name === 'w:r');
  const pict = run?.elements?.find((el) => el.name === 'w:pict');

  // if there is no pict, then process as a paragraph or list.
  if (!pict) {
    return { nodes: [], consumed: 0 };
  }

  const node = pict;
  const shape = node.elements?.find((el) => el.name === 'v:shape');
  const shapetype = node.elements?.find((el) => el.name === 'v:shapetype');
  const group = node.elements?.find((el) => el.name === 'v:group');

  // such a case probably shouldn't exist.
  if (!shape && !group) {
    return { nodes: [], consumed: 1 };
  }

  let result = null;

  const isGroup = group && !shape;

  if (isGroup) {
    // there should be a group of shapes being processed here (skip for now).
    result = null;
  } else {
    const textbox = shape.elements?.find((el) => el.name === 'v:textbox');

    // process shapes with textbox.
    if (textbox) {
      result = handleShapTextboxImport({
        node,
        pNode,
        shape, 
        params,
      });
    }
  }

  return { nodes: result ? [result] : [], consumed: 1 };
};

export function handleShapTextboxImport({
  node, 
  pNode,
  shape, 
  params,
}) {
  const schemaAttrs = {};
  const schemaTextboxAttrs = {};

  const shapeAttrs = shape.attributes || {};

  schemaAttrs.attributes = shapeAttrs;

  if (shapeAttrs.fillcolor) {
    schemaAttrs.fillcolor = shapeAttrs.fillcolor;
  }
  if (shapeAttrs.style) {
    schemaAttrs.style = shapeAttrs.style;
  }

  const textbox = shape.elements?.find((el) => el.name === 'v:textbox');
  const wrap = shape.elements?.find((el) => el.name === 'w10:wrap');

  if (wrap?.attributes) {
    schemaAttrs.wrapAttributes = wrap.attributes;
  }

  if (textbox?.attributes) {
    schemaTextboxAttrs.attributes = textbox.attributes;
  }

  const textboxContent = textbox?.elements?.find((el) => el.name === 'w:txbxContent');
  const textboxContentElems = textboxContent?.elements || [];

  const content = textboxContentElems.map((elem) => handleParagraphNode({ 
    nodes: [elem],
    docx: params.docx,
    nodeListHandler: defaultNodeListHandler(),
  }));
  const contentNodes = content.reduce((acc, current) => (
    [...acc, ...current.nodes]
  ), []);

  const shapeTextbox = {
    type: 'shapeTextbox',
    attrs: schemaTextboxAttrs,
    content: contentNodes,
  };

  const shapeContainer = {
    type: 'shapeContainer',
    attrs: schemaAttrs,
    content: [shapeTextbox],
  };

  return shapeContainer;
}

export const pictNodeHandlerEntity = {
  handlerName: 'handlePictNode',
  handler: handlePictNode,
};
