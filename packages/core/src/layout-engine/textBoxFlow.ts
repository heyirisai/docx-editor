import { isFloatingWrapType, isWrapNone } from '../docx/wrapTypes';
import type { TextBoxBlock } from './types';

export type TextBoxFlowAttrs = Pick<TextBoxBlock, 'displayMode' | 'wrapType'>;

export function isFloatingTextBoxBlock(block: TextBoxFlowAttrs): boolean {
  return block.displayMode === 'float' || isFloatingWrapType(block.wrapType);
}

export function floatingTextBoxWrapsText(block: TextBoxFlowAttrs): boolean {
  return (
    isFloatingTextBoxBlock(block) &&
    !isWrapNone(block.wrapType) &&
    block.wrapType !== 'topAndBottom'
  );
}
