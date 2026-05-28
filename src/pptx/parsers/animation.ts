import type { ParsedNode } from '../../core/types.js';
import type { Animation } from '../types.js';
import { findChild } from './utils.js';

export function parseAnimations(node: ParsedNode): Animation[] {
  const timing = findChild(node, 'p:timing');
  if (!timing) return [];

  const tnLst = findChild(timing, 'p:tnLst');
  if (!tnLst) return [];

  return parseAnimationList(tnLst);
}

function parseAnimationList(node: ParsedNode): Animation[] {
  const animations: Animation[] = [];

  for (const child of node.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'p:par') {
      const anim = parseParallelAnimation(child);
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:seq') {
      const anim = parseSequenceAnimation(child);
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:set') {
      const anim = parseSetAnimation(child, 'onClick');
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:anim') {
      const anim = parseAnimNode(child, 'onClick');
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:animEffect') {
      const anim = parseAnimEffectNode(child, 'onClick');
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:animMotion') {
      const anim = parseAnimMotionNode(child, 'onClick');
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:animScale') {
      const anim = parseAnimScaleNode(child, 'onClick');
      if (anim) animations.push(anim);
    } else if (child.tag === 'p:animRotation') {
      const anim = parseAnimRotationNode(child, 'onClick');
      if (anim) animations.push(anim);
    }
  }

  return animations;
}

function parseParallelAnimation(node: ParsedNode): Animation | undefined {
  const cTn = findChild(node, 'p:cTn');
  if (!cTn) return undefined;

  const trigger = parseTrigger(cTn);
  const duration = parseDuration(cTn);

  const childTnLst = findChild(cTn, 'p:childTnLst');
  if (!childTnLst) return undefined;

  const children = parseAnimationList(childTnLst);

  return {
    trigger,
    type: 'parallel',
    duration,
    children,
  };
}

function parseSequenceAnimation(node: ParsedNode): Animation | undefined {
  const cTn = findChild(node, 'p:cTn');
  if (!cTn) return undefined;

  const trigger = parseTrigger(cTn);
  const duration = parseDuration(cTn);

  const childTnLst = findChild(cTn, 'p:childTnLst');
  if (!childTnLst) return undefined;

  const children = parseAnimationList(childTnLst);

  return {
    trigger,
    type: 'sequence',
    duration,
    children,
  };
}

function parseTrigger(cTn: ParsedNode): Animation['trigger'] {
  const triggerType = cTn.attrs['triggerType'];
  const masterRel = cTn.attrs['masterRel'];

  if (triggerType === 'onClick') return 'onClick';
  if (triggerType === 'withPrev') return 'withPrevious';
  if (triggerType === 'afterPrev') return 'afterPrevious';
  if (triggerType === 'onLoad') return 'onLoad';

  // Default based on master relationship
  if (masterRel === 'onClick') return 'onClick';
  if (masterRel === 'withPrev') return 'withPrevious';
  if (masterRel === 'afterPrev') return 'afterPrevious';

  return 'onClick';
}

function parseDuration(cTn: ParsedNode): number | undefined {
  const dur = cTn.attrs['dur'];
  if (dur) return parseInt(dur, 10);
  return undefined;
}

export function parseEffectAnimations(node: ParsedNode): Animation[] {
  const timing = findChild(node, 'p:timing');
  if (!timing) return [];

  const tnLst = findChild(timing, 'p:tnLst');
  if (!tnLst) return [];

  const animations: Animation[] = [];

  for (const child of tnLst.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'p:par') {
      const anim = parseEffectNode(child);
      if (anim) animations.push(anim);
    }
  }

  return animations;
}

function parseEffectNode(node: ParsedNode): Animation | undefined {
  const cTn = findChild(node, 'p:cTn');
  if (!cTn) return undefined;

  const trigger = parseTrigger(cTn);
  const duration = parseDuration(cTn);
  const delay = parseDelay(cTn);

  const childTnLst = findChild(cTn, 'p:childTnLst');
  if (!childTnLst) return undefined;

  const children: Animation[] = [];

  for (const child of childTnLst.children) {
    if (typeof child === 'string') continue;

    if (child.tag === 'p:set') {
      const setAnim = parseSetAnimation(child, trigger);
      if (setAnim) children.push(setAnim);
    } else if (child.tag === 'p:anim') {
      const anim = parseAnimNode(child, trigger);
      if (anim) children.push(anim);
    } else if (child.tag === 'p:animEffect') {
      const effect = parseAnimEffectNode(child, trigger);
      if (effect) children.push(effect);
    } else if (child.tag === 'p:animMotion') {
      const motion = parseAnimMotionNode(child, trigger);
      if (motion) children.push(motion);
    } else if (child.tag === 'p:animScale') {
      const scale = parseAnimScaleNode(child, trigger);
      if (scale) children.push(scale);
    } else if (child.tag === 'p:animRotation') {
      const rotation = parseAnimRotationNode(child, trigger);
      if (rotation) children.push(rotation);
    }
  }

  return {
    trigger,
    type: 'effect',
    duration,
    delay,
    children,
  };
}

function parseDelay(cTn: ParsedNode): number | undefined {
  const delay = cTn.attrs['delay'];
  if (delay) return parseInt(delay, 10);
  return undefined;
}

function parseSetAnimation(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);

  return {
    trigger,
    type: 'set',
    duration,
    shapeId,
  };
}

function parseAnimNode(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);

  return {
    trigger,
    type: 'animate',
    duration,
    shapeId,
  };
}

function parseAnimEffectNode(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);
  const transition = node.attrs['transition'] || 'in';

  return {
    trigger,
    type: 'effect',
    duration,
    shapeId,
    direction: transition,
  };
}

function parseAnimMotionNode(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);

  return {
    trigger,
    type: 'motion',
    duration,
    shapeId,
  };
}

function parseAnimScaleNode(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);

  return {
    trigger,
    type: 'scale',
    duration,
    shapeId,
  };
}

function parseAnimRotationNode(node: ParsedNode, trigger: Animation['trigger']): Animation | undefined {
  const cBhvr = findChild(node, 'p:cBhvr');
  if (!cBhvr) return undefined;

  const cTn = findChild(cBhvr, 'p:cTn');
  const duration = cTn ? parseDuration(cTn) : undefined;
  const shapeId = parseTargetShape(cBhvr);

  return {
    trigger,
    type: 'rotation',
    duration,
    shapeId,
  };
}

function parseTargetShape(cBhvr: ParsedNode): string | undefined {
  const tgtEl = findChild(cBhvr, 'p:tgtEl');
  if (!tgtEl) return undefined;

  const spTgt = findChild(tgtEl, 'p:spTgt');
  if (!spTgt) return undefined;

  return spTgt.attrs['spid'];
}
