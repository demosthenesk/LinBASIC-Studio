# LinBASIC Studio - lbvc: the order in which VB6 evaluates the parts of an expression
# Copyright (C) 2026 Demosthenes Koptsis
# SPDX-License-Identifier: GPL-3.0-or-later
"""VB6 does not evaluate an expression simply from left to right, and the order depends on how the program
is compiled. This model was measured on VB6 SP6 (tests/order: about 18,000 observations, P-code and native
code with every optimization setting) and predicts every case of those experiments:

* every mode: the arguments of a call except the last are evaluated first, into temporaries, left to right;
  the last one just before the call. A floating-point operator (Single/Double/Date, `/`, `^`, a String with
  a number; not Currency or Variant) evaluates the calls of its right operand first and its left operand last.
* P-code (and the IDE): everything else from left to right.
* native code: Currency + - * from right to left.
* native Fast/Small code, in a procedure without any On Error statement: the operands of integer operators
  (Integer/Long/Byte/Boolean, And/Or, comparisons of integers) are evaluated heavier first (Sethi-Ullman
  number, ties: the right one first); never around object methods, comparisons of floats/Currency or
  narrowing CInt/CLng. No Optimization never reorders them.
* "Remove Integer Overflow Checks" (OverflowCheck=-1) with Fast/Small: chains of + or * are flattened and
  their operands evaluated heavier first (ties: from the right); Currency follows the integer rule (only *
  swaps its operands in a procedure with On Error).

The tree is made of tuples:
    ('leaf', ev, node, method)            a call that is an event of its own (ev: its index)
    ('const', node) / ('var', node)       no event
    ('neg', e) / ('not', e)
    ('call', name, args, node, ev, method) a call with arguments; ev: its own event (or None)
    ('bin', op, l, r)                     op one of + - * / \\ Mod ^ & = Like And Or Xor
The types of leaves, constants, variables and calls come from the callback tyof(node).
"""

FLOAT = {'Single', 'Double', 'Date'}
RANK = ['Byte', 'Boolean', 'Integer', 'Long', 'Single', 'Double', 'Currency', 'Date']
CONV = {'cint': 'Integer', 'clng': 'Long', 'cdbl': 'Double', 'cdate': 'Date'}
TRANSPARENT = ('abs', 'cdbl', 'clng', 'cint', 'cdate', 'len')
NARROW = ('Integer', 'Byte', 'Boolean')
WIDEN = {('Integer', 'Long'), ('Byte', 'Integer'), ('Byte', 'Long'), ('Boolean', 'Integer'), ('Boolean', 'Long'),
         ('Integer', 'Double'), ('Long', 'Double'), ('Single', 'Double'), ('Byte', 'Double')}


class Model:
    def __init__(self, mode, handler, tyof):
        """mode: 'pcode', 'noopt', 'fast' or 'adv' (Fast/Small with the overflow checks removed);
        handler: the procedure has an On Error statement"""
        self.mode, self.handler, self.tyof = mode, handler, tyof
        self.su = mode in ('fast', 'adv') and not handler
        self.cache = {}

    # ------------------------------------------------------------ types
    def typ(self, e):
        k = e[0]
        if k in ('leaf', 'const', 'var'):
            return self.tyof(e[2] if k == 'leaf' else e[1])
        if k == 'call':
            n = e[1]
            if n in CONV:
                return CONV[n]
            if n == 'abs' and e[2]:
                return self.typ(e[2][0])
            return self.tyof(e[3])
        if k in ('neg', 'not'):
            return self.typ(e[1])
        key = id(e)
        if key in self.cache:
            return self.cache[key][1]
        op, a, b = e[1], self.typ(e[2]), self.typ(e[3])
        if 'Variant' in (a, b):
            t = 'Variant'
        elif op in ('=', 'Like'):
            t = 'Boolean'
        elif op == '&':
            t = 'String'
        elif op in ('/', '^'):
            t = 'Double' if 'Currency' not in (a, b) or op == '^' else 'Currency'
        elif 'String' in (a, b):
            t = 'Double'
        elif op in ('And', 'Or', 'Xor', '\\', 'Mod'):
            t = 'Long'
        elif 'Date' in (a, b):
            t = 'Date'
        elif 'Currency' in (a, b):
            t = 'Currency'
        else:
            t = max(a, b, key=lambda x: RANK.index(x) if x in RANK else -1)
        self.cache[key] = (e, t)
        return t

    def is_float_op(self, e):
        op, a, b = e[1], self.typ(e[2]), self.typ(e[3])
        if 'Variant' in (a, b) or op in ('&', 'Like', 'And', 'Or', 'Xor', '\\', 'Mod'):
            return False
        if op in ('/', '^'):
            return True
        if 'Currency' in (a, b):
            return False
        if 'String' in (a, b):
            return not (a == b == 'String')
        return bool({a, b} & FLOAT)

    def kind(self, e):
        op, a, b = e[1], self.typ(e[2]), self.typ(e[3])
        if self.is_float_op(e):
            return 'float'
        if 'Variant' in (a, b) or 'String' in (a, b) or op in ('&', 'Like'):
            return 'other'
        if 'Currency' in (a, b) and op in ('+', '-', '*'):
            return 'cy'
        return 'int'

    # ------------------------------------------------------------ shape
    def has_method(self, e):
        k = e[0]
        if k == 'leaf':
            return e[3]
        if k == 'call':
            return e[5] or any(self.has_method(a) for a in e[2])
        if k in ('neg', 'not'):
            return self.has_method(e[1])
        if k == 'bin':
            return self.has_method(e[2]) or self.has_method(e[3])
        return False

    def is_conv(self, e):
        return e[0] == 'call' and e[1] in CONV and len(e[2]) == 1 and self.typ(e[2][0]) != CONV[e[1]]

    def weight(self, e):
        k = e[0]
        if k in ('leaf', 'const'):
            return 1
        if k == 'var':                              # (VB6 exp1: Long k + Integer call + Integer call)
            return 2
        if k in ('neg', 'not'):
            return self.weight(e[1])
        if k == 'call':
            n = e[1]
            if n in CONV and len(e[2]) == 1:
                return self.weight(e[2][0]) + (1 if self.is_conv(e) else 0)
            if n in ('abs', 'len') and len(e[2]) == 1:
                return self.weight(e[2][0])
            if n == '()':                          # an element of an array: the index, then the load
                return 1 + max([self.weight(a) for a in e[2]] or [0])
            if any(self.leaves(a) for a in e[2]):  # a call with calls in its arguments (VB6 exp1)
                return 1 + max(self.weight(a) for a in e[2])
            return 1
        l, r = self.operand_weights(e)
        return l + 1 if l == r else max(l, r)

    def operand_weights(self, e):
        l, r = self.weight(e[2]), self.weight(e[3])
        ta, tb = self.typ(e[2]), self.typ(e[3])
        wide = 'Long' in (ta, tb) and self.kind(e) == 'int'
        if wide and ta in NARROW and not self.is_conv(e[2]):
            l += 1                                  # an Integer made Long: one more step
        if wide and tb in NARROW and not self.is_conv(e[3]):
            r += 1
        return l, r

    def barrier(self, e):
        """an operand computed by a runtime helper: the optimizer does not reorder around it"""
        k = e[0]
        if k == 'call' and e[1] in ('cint', 'clng', 'cdbl') and len(e[2]) == 1:
            t, want = self.typ(e[2][0]), CONV[e[1]]
            return not (t == want or (t, want) in WIDEN)
        if k == 'bin' and e[1] == '=' and (self.kind(e) == 'float' or 'Currency' in (self.typ(e[2]), self.typ(e[3]))):
            return True
        return False

    def flatten(self, e, op):
        if e[0] == 'bin' and e[1] == op and self.kind(e) in ('int', 'cy'):
            return self.flatten(e[2], op) + self.flatten(e[3], op)
        return [e]

    # ------------------------------------------------------------ order
    def leaves(self, e):
        """the events of e from left to right"""
        k = e[0]
        if k == 'leaf':
            return [e[1]]
        if k == 'call':
            return [x for a in e[2] for x in self.leaves(a)] + ([e[4]] if e[4] is not None else [])
        if k in ('neg', 'not'):
            return self.leaves(e[1])
        if k == 'bin':
            return self.leaves(e[2]) + self.leaves(e[3])
        return []

    def events(self, e):
        p1, p2 = self.order(e)
        return p1 + p2

    def order(self, e):
        """(phase 1, phase 2): phase 1 is evaluated first (into temporaries), then phase 2"""
        k = e[0]
        if k == 'leaf':
            return [], [e[1]]
        if k in ('const', 'var'):
            return [], []
        if k in ('neg', 'not'):
            return self.order(e[1])
        if k == 'call':
            if e[1] in TRANSPARENT and len(e[2]) == 1 and e[4] is None:
                return self.order(e[2][0])
            p1, p2 = [], []
            for i, a in enumerate(e[2]):
                if i < len(e[2]) - 1:
                    p1 += self.leaves(a)
                else:
                    a1, a2 = self.order(a)
                    p1 += a1
                    p2 += a2
            if e[4] is not None:
                p2.append(e[4])
            return p1, p2
        kd = self.kind(e)
        if kd == 'float':
            l1, l2 = self.order(e[2])
            return l1 + self.leaves(e[3]), l2
        native = self.mode != 'pcode'
        L, R = e[2], e[3]
        if kd == 'cy' and native and not (self.mode == 'adv' and not self.handler):
            if self.mode == 'adv':                  # with a handler only * swaps its operands
                return self.pair(R, L) if e[1] == '*' else self.pair(L, R)
            return self.pair(R, L)                  # Currency + - *: right to left
        if kd == 'cy' and native and self.has_method(e):
            return self.pair(R, L) if e[1] == '*' else self.pair(L, R)
        if kd in ('int', 'cy') and self.su and not self.has_method(e) and (
                not (self.barrier(L) or self.barrier(R)) if kd == 'int'
                else not ({self.typ(L), self.typ(R)} & FLOAT)):
            if self.mode == 'adv' and e[1] in ('+', '*') and \
                    self.typ(e) in ('Integer', 'Long', 'Byte', 'Boolean', 'Currency') and \
                    not (self.typ(e) == 'Boolean' and any(x[0] == 'const' for x in self.flatten(e, e[1]))):
                ops = self.flatten(e, e[1])
                if len(ops) > 2:
                    idx = sorted(range(len(ops)), key=lambda i: (-self.weight(ops[i]), -i))
                    p1, p2 = [], []
                    for i in idx:
                        a1, a2 = self.order(ops[i])
                        p1 += a1
                        p2 += a2
                    return p1, p2
            l, r = self.operand_weights(e)
            return self.pair(L, R) if l > r else self.pair(R, L)
        return self.pair(L, R)

    def pair(self, first, second):
        a1, a2 = self.order(first)
        b1, b2 = self.order(second)
        return a1 + b1, a2 + b2
