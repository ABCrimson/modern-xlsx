/**
 * Formula AST reference walker shared by the rewriter and shared-formula
 * expansion. Both rebuild a tree the same way and differ only in what they do
 * at the reference leaves; this module owns the tree shape once.
 *
 * @module formula/walk
 */

import type {
  ASTNode,
  BinaryOpNode,
  CellRefNode,
  ColumnRangeNode,
  PercentNode,
  RangeNode,
  RowRangeNode,
} from './parser.js';

/** What to do at each reference node; every composite node is rebuilt by {@link mapRefs}. */
export interface RefMappers {
  cellRef(node: CellRefNode): ASTNode;
  range(node: RangeNode): ASTNode;
  columnRange(node: ColumnRangeNode): ASTNode;
  rowRange(node: RowRangeNode): ASTNode;
}

/**
 * Rebuild an AST node (a structural copy: composite nodes are new, unchanged
 * leaves are shared with the input), replacing each reference node with the
 * result of the matching mapper. Composite nodes (functions, operators,
 * percent, arrays) are rebuilt with their children mapped; other leaves
 * (numbers, strings, booleans, errors, names, missing arguments, structured
 * references) are returned as they are. A left-folded chain of binary
 * operators and `%` (`A1+B1+C1`, `A1%%`) is walked with a loop, so a chain
 * thousands of levels deep never exhausts the call stack.
 */
export function mapRefs(node: ASTNode, leaf: RefMappers): ASTNode {
  if (node.type === 'binary_op' || node.type === 'percent') return mapChain(node, leaf);
  switch (node.type) {
    case 'cell_ref':
      return leaf.cellRef(node);

    case 'range':
      return leaf.range(node);

    case 'column_range':
      return leaf.columnRange(node);

    case 'row_range':
      return leaf.rowRange(node);

    case 'function':
      return {
        type: 'function',
        name: node.name,
        args: node.args.map((arg) => mapRefs(arg, leaf)),
      };

    case 'unary_op':
      return {
        type: 'unary_op',
        op: node.op,
        operand: mapRefs(node.operand, leaf),
      };

    case 'array':
      return {
        type: 'array',
        rows: node.rows.map((row) => row.map((cell) => mapRefs(cell, leaf))),
      };

    // Leaf nodes without position-relative references
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
    case 'name':
    case 'missing':
    case 'structured_ref':
      return node;
  }
}

/**
 * {@link mapRefs} for a binary operator (the reference operators `:`, ` `
 * and `,` included) or `%`: descend the left spine with a loop, map the
 * innermost operand, then rebuild each link outwards, mapping its right
 * operand in the same left-then-right order as a recursive walk.
 */
function mapChain(node: BinaryOpNode | PercentNode, leaf: RefMappers): ASTNode {
  const links: (BinaryOpNode | PercentNode)[] = [];
  let innermost: ASTNode = node;
  while (innermost.type === 'binary_op' || innermost.type === 'percent') {
    links.push(innermost);
    innermost = innermost.type === 'percent' ? innermost.operand : innermost.left;
  }
  let acc = mapRefs(innermost, leaf);
  for (const link of links.toReversed()) {
    acc =
      link.type === 'percent'
        ? { type: 'percent', operand: acc }
        : { type: 'binary_op', op: link.op, left: acc, right: mapRefs(link.right, leaf) };
  }
  return acc;
}
