import type { ParsedNode } from '../../core/types.js';
import type { Transition, Animation } from '../types.js';

export function serializeTransition(transition: Transition): ParsedNode {
  const speed = transition.duration && transition.duration <= 500 ? 'fast' :
    transition.duration && transition.duration <= 1000 ? 'med' : 'slow';

  const transitionChildren: ParsedNode[] = [];
  const typeMap: Record<string, string> = {
    fade: 'p:fade',
    blinds: 'p:blinds',
    checker: 'p:checker',
    circle: 'p:circle',
    dissolve: 'p:dissolve',
    newsflash: 'p:newsflash',
    plus: 'p:plus',
    pull: 'p:pull',
    push: 'p:push',
    random: 'p:random',
    split: 'p:split',
    strips: 'p:strips',
    wedge: 'p:wedge',
    wheel: 'p:wheel',
    wipe: 'p:wipe',
    zoom: 'p:zoom',
  };

  const tag = typeMap[transition.type];
  if (tag) {
    transitionChildren.push({ tag, attrs: {} as Record<string, string>, children: [] });
  }

  return {
    tag: 'p:transition',
    attrs: { spd: speed },
    children: transitionChildren,
  };
}

export function serializeTiming(animations: Animation[]): ParsedNode {
  return {
    tag: 'p:timing',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:tnLst',
        attrs: {} as Record<string, string>,
        children: animations.map(anim => serializeAnimation(anim)),
      },
    ],
  };
}

function serializeAnimation(anim: Animation): ParsedNode {
  const triggerType = anim.trigger === 'onClick' ? 'onClick' :
    anim.trigger === 'withPrevious' ? 'withPrev' :
    anim.trigger === 'afterPrevious' ? 'afterPrev' : 'onLoad';

  const cTnAttrs: Record<string, string> = {
    id: '1',
    presetID: '0',
    presetClass: 'entr',
    triggerType,
  };

  if (anim.duration) cTnAttrs.dur = String(anim.duration);
  if (anim.delay) cTnAttrs.delay = String(anim.delay);

  const childTnLst: ParsedNode[] = [];
  if (anim.children) {
    for (const child of anim.children) {
      childTnLst.push(serializeAnimationNode(child));
    }
  }

  return {
    tag: 'p:par',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:cTn',
        attrs: cTnAttrs,
        children: [
          {
            tag: 'p:childTnLst',
            attrs: {} as Record<string, string>,
            children: childTnLst,
          },
        ],
      },
    ],
  };
}

function serializeAnimationNode(anim: Animation): ParsedNode {
  if (anim.type === 'set') {
    return serializeSetAnimation(anim);
  } else if (anim.type === 'animate') {
    return serializeAnimNode(anim);
  } else if (anim.type === 'effect') {
    return serializeAnimEffectNode(anim);
  }

  return {
    tag: 'p:par',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:cTn',
        attrs: { id: '1' },
        children: [],
      },
    ],
  };
}

function serializeSetAnimation(anim: Animation): ParsedNode {
  const cBhvrChildren: ParsedNode[] = [];

  if (anim.shapeId) {
    cBhvrChildren.push({
      tag: 'p:tgtEl',
      attrs: {} as Record<string, string>,
      children: [
        { tag: 'p:spTgt', attrs: { spid: anim.shapeId }, children: [] },
      ],
    });
  }

  return {
    tag: 'p:set',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:cBhvr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'p:cTn',
            attrs: { id: '1', ...(anim.duration ? { dur: String(anim.duration) } : {}) },
            children: [],
          },
          ...cBhvrChildren,
        ],
      },
    ],
  };
}

function serializeAnimNode(anim: Animation): ParsedNode {
  const cBhvrChildren: ParsedNode[] = [];

  if (anim.shapeId) {
    cBhvrChildren.push({
      tag: 'p:tgtEl',
      attrs: {} as Record<string, string>,
      children: [
        { tag: 'p:spTgt', attrs: { spid: anim.shapeId }, children: [] },
      ],
    });
  }

  return {
    tag: 'p:anim',
    attrs: {} as Record<string, string>,
    children: [
      {
        tag: 'p:cBhvr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'p:cTn',
            attrs: { id: '1', ...(anim.duration ? { dur: String(anim.duration) } : {}) },
            children: [],
          },
          ...cBhvrChildren,
        ],
      },
    ],
  };
}

function serializeAnimEffectNode(anim: Animation): ParsedNode {
  const cBhvrChildren: ParsedNode[] = [];

  if (anim.shapeId) {
    cBhvrChildren.push({
      tag: 'p:tgtEl',
      attrs: {} as Record<string, string>,
      children: [
        { tag: 'p:spTgt', attrs: { spid: anim.shapeId }, children: [] },
      ],
    });
  }

  const attrs: Record<string, string> = {};
  if (anim.direction) attrs.transition = anim.direction;

  return {
    tag: 'p:animEffect',
    attrs,
    children: [
      {
        tag: 'p:cBhvr',
        attrs: {} as Record<string, string>,
        children: [
          {
            tag: 'p:cTn',
            attrs: { id: '1', ...(anim.duration ? { dur: String(anim.duration) } : {}) },
            children: [],
          },
          ...cBhvrChildren,
        ],
      },
    ],
  };
}
