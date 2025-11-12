/**
 * Copyright (c) 2011 Lorenz Schori <lo@znerol.ch>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

/**
 * @file:   Implementation of Myers linear space longest common subsequence
 *          algorithm.
 * @see:
 * * http://dx.doi.org/10.1007/BF01840446
 * * http://citeseer.ist.psu.edu/viewdoc/summary?doi=10.1.1.4.6927
 *
 * @module  lcs
 */

(function (window, undefined) {
    /**
     * Create a new instance of the LCS implementation.
     *
     * @param a     The first sequence
     * @param b     The second sequence
     *
     * @constructor
     */
    function LCS(a, b) {
        this.a = a;
        this.b = b;
    }


    /**
     * Returns true if the sequence members a and b are equal. Override this
     * method if your sequences contain special things.
     */
    LCS.prototype.equals = function(a, b) {
        return (a === b);
    };


    /**
     * Compute longest common subsequence using myers divide & conquer linear
     * space algorithm.
     *
     * Call a callback for each snake which is part of the longest common
     * subsequence.
     *
     * This algorithm works with strings and arrays. In order to modify the
     * equality-test, just override the equals(a, b) method on the LCS
     * object.
     *
     * @param callback  A function(x, y) called for A[x] and B[y] for symbols
     *                  taking part in the LCS.
     * @param T         Context object bound to "this" when the callback is
     *                  invoked.
     * @param limit     A Limit instance constraining the window of operation to
     *                  the given limit. If undefined the algorithm will iterate
     *                  over the whole sequences a and b.
     */
    LCS.prototype.compute = function(callback, T, limit) {
        var midleft = new KPoint(),
            midright = new KPoint(),
            d;

        if (typeof limit === 'undefined') {
            limit = this.defaultLimit();
        }

        // Return if there is nothing left
        if (limit.N <= 0 && limit.M <= 0) {
            return 0;
        }

        // Callback for each right-edge when M is zero and return number of
        // edit script operations.
        if (limit.N > 0 && limit.M === 0) {
            midleft.set(0, 0).translate(limit.left);
            midright.set(1, 1).translate(limit.left);
            for (d = 0; d < limit.N; d++) {
                callback.call(T, midleft, midright);
                midleft.moveright();
                midright.moveright();
            }
            return d;
        }

        // Callback for each down-edge when N is zero and return number of edit
        // script operations.
        if (limit.N === 0 && limit.M > 0) {
            midleft.set(0, 0).translate(limit.left);
            midright.set(0, -1).translate(limit.left);
            for (d = 0; d < limit.M; d++) {
                callback.call(T, midleft, midright);
                midleft.movedown();
                midright.movedown();
            }
            return d;
        }

        // Find the middle snake and store the result in midleft and midright
        d = this.middleSnake(midleft, midright, limit);

        if (d === 0) {
            // No single insert / delete operation was identified by the middle
            // snake algorithm, this means that all the symbols between left and
            // right are equal -> one straight diagonal on k=0
            if (!limit.left.equal(limit.right)) {
                callback.call(T, limit.left, limit.right);
            }
        }
        else if (d === 1) {
            // Middle-snake algorithm identified exactly one operation. Report
            // the involved snake(s) to the caller.
            if (!limit.left.equal(midleft)) {
                callback.call(T, limit.left, midleft);
            }

            if (!midleft.equal(midright)) {
                callback.call(T, midleft, midright);
            }

            if (!midright.equal(limit.right)) {
                callback.call(T, midright, limit.right);
            }
        }
        else {
            // Recurse if the middle-snake algorithm encountered more than one
            // operation.
            if (!limit.left.equal(midleft)) {
                this.compute(callback, T, new Limit(limit.left, midleft));
            }

            if (!midleft.equal(midright)) {
                callback.call(T, midleft, midright);
            }

            if (!midright.equal(limit.right)) {
                this.compute(callback, T, new Limit(midright, limit.right));
            }
        }

        return d;
    };


    /**
     * Call a callback for each symbol which is part of the longest common
     * subsequence between A and B.
     *
     * Given that the two sequences A and B were supplied to the LCS
     * constructor, invoke the callback for each pair A[x], B[y] which is part
     * of the longest common subsequence of A and B.
     *
     * This algorithm works with strings and arrays. In order to modify the
     * equality-test, just override the equals(a, b) method on the LCS
     * object.
     *
     * Usage:
     * <code>
     * var lcs = [];
     * var A = 'abcabba';
     * var B = 'cbabac';
     * var l = new LCS(A, B);
     * l.forEachCommonSymbol(function(x, y) {
 *     lcs.push(A[x]);
 * });
     * console.log(lcs);
     * // -> [ 'c', 'a', 'b', 'a' ]
     * </code>
     *
     * @param callback  A function(x, y) called for A[x] and B[y] for symbols
     *                  taking part in the LCS.
     * @param T         Context object bound to "this" when the callback is
     *                  invoked.
     */
    LCS.prototype.forEachCommonSymbol = function(callback, T) {
        return this.compute(function(left, right) {
            this.forEachPositionInSnake(left, right, callback, T);
        }, this);
    };


    /**
     * Internal use. Compute new values for the next head on the given k-line
     * in forward direction by examining the results of previous calculations
     * in V in the neighborhood of the k-line k.
     *
     * @param head  (Output) Reference to a KPoint which will be populated
     *              with the new values
     * @param k     (In) Current k-line
     * @param kmin  (In) Lowest k-line in current d-round
     * @param kmax  (In) Highest k-line in current d-round
     * @param limit (In) Current lcs search limits (left, right, N, M, delta, dmax)
     * @param V     (In-/Out) Vector containing the results of previous
     *              calculations. This vector gets updated automatically by
     *              nextSnakeHeadForward method.
     */
    LCS.prototype.nextSnakeHeadForward = function(head, k, kmin, kmax, limit, V) {
        var k0, x, bx, by, n;

        // Determine the preceeding snake head. Pick the one whose furthest
        // reaching x value is greatest.
        if (k === kmin || (k !== kmax && V[k-1] < V[k+1])) {
            // Furthest reaching snake is above (k+1), move down.
            k0 = k+1;
            x = V[k0];
        }
        else {
            // Furthest reaching snake is left (k-1), move right.
            k0 = k-1;
            x = V[k0] + 1;
        }

        // Follow the diagonal as long as there are common values in a and b.
        bx = limit.left.x;
        by = bx - (limit.left.k + k);
        n = Math.min(limit.N, limit.M + k);
        while (x < n && this.equals(this.a[bx + x], this.b[by + x])) {
            x++;
        }

        // Store x value of snake head after traversing the diagonal in forward
        // direction.
        head.set(x, k).translate(limit.left);

        // Memozie furthest reaching x for k
        V[k] = x;

        // Return k-value of preceeding snake head
        return k0;
    };


    /**
     * Internal use. Compute new values for the next head on the given k-line
     * in reverse direction by examining the results of previous calculations
     * in V in the neighborhood of the k-line k.
     *
     * @param head  (Output) Reference to a KPoint which will be populated
     *              with the new values
     * @param k     (In) Current k-line
     * @param kmin  (In) Lowest k-line in current d-round
     * @param kmax  (In) Highest k-line in current d-round
     * @param limit (In) Current lcs search limits (left, right, N, M, delta, dmax)
     * @param V     (In-/Out) Vector containing the results of previous
     *              calculations. This vector gets updated automatically by
     *              nextSnakeHeadForward method.
     */
    LCS.prototype.nextSnakeHeadBackward = function(head, k, kmin, kmax, limit, V) {
        var k0, x, bx, by, n;

        // Determine the preceeding snake head. Pick the one whose furthest
        // reaching x value is greatest.
        if (k === kmax || (k !== kmin && V[k-1] < V[k+1])) {
            // Furthest reaching snake is underneath (k-1), move up.
            k0 = k-1;
            x = V[k0];
        }
        else {
            // Furthest reaching snake is left (k-1), move right.
            k0 = k+1;
            x = V[k0]-1;
        }

        // Store x value of snake head before traversing the diagonal in
        // reverse direction.
        head.set(x, k).translate(limit.left);

        // Follow the diagonal as long as there are common values in a and b.
        bx = limit.left.x - 1;
        by = bx - (limit.left.k + k);
        n = Math.max(k, 0);
        while (x > n && this.equals(this.a[bx + x], this.b[by + x])) {
            x--;
        }

        // Memozie furthest reaching x for k
        V[k] = x;

        // Return k-value of preceeding snake head
        return k0;
    };


    /**
     * Internal use. Find the middle snake and set lefthead to the left end and
     * righthead to the right end.
     *
     * @param lefthead  (Output) A reference to a KPoint which will be
     *                  populated with the values corresponding to the left end
     *                  of the middle snake.
     * @param righthead (Output) A reference to a KPoint which will be
     *                  populated with the values corresponding to the right
     *                  end of the middle snake.
     * @param limit     (In) Current lcs search limits (left, right, N, M, delta, dmax)
     *
     * @returns         d, number of edit script operations encountered within
     *                  the given limit
     */
    LCS.prototype.middleSnake = function (lefthead, righthead, limit) {
        var d, k, head, k0;
        var delta = limit.delta;
        var dmax = Math.ceil(limit.dmax / 2);
        var checkBwSnake = (delta % 2 === 0);
        var Vf = {};
        var Vb = {};

        Vf[1] = 0;
        Vb[delta-1] = limit.N;
        for (d = 0; d <= dmax; d++) {
            for (k = -d; k <= d; k+=2) {
                k0 = this.nextSnakeHeadForward(righthead, k, -d, d, limit, Vf);

                // check for overlap
                if (!checkBwSnake && k >= -d-1+delta && k <= d-1+delta) {
                    if (Vf[k] >= Vb[k]) {
                        // righthead already contains the right stuff, now set
                        // the lefthead to the values of the last k-line.
                        lefthead.set(Vf[k0], k0).translate(limit.left);
                        // return the number of edit script operations
                        return 2 * d - 1;
                    }
                }
            }

            for (k = -d+delta; k <= d+delta; k+=2) {
                k0 = this.nextSnakeHeadBackward(lefthead, k, -d+delta, d+delta, limit, Vb);

                // check for overlap
                if (checkBwSnake && k >= -d && k <= d) {
                    if (Vf[k] >= Vb[k]) {
                        // lefthead already contains the right stuff, now set
                        // the righthead to the values of the last k-line.
                        righthead.set(Vb[k0], k0).translate(limit.left);
                        // return the number of edit script operations
                        return 2 * d;
                    }
                }
            }
        }
    };


    /**
     * Return the default limit spanning the whole input
     */
    LCS.prototype.defaultLimit = function() {
        return new Limit(
            new KPoint(0,0),
            new KPoint(this.a.length, this.a.length - this.b.length));
    };


    /**
     * Invokes a function for each position in the snake between the left and
     * the right snake head.
     *
     * @param left      Left KPoint
     * @param right     Right KPoint
     * @param callback  Callback of the form function(x, y)
     * @param T         Context object bound to "this" when the callback is
     *                  invoked.
     */
    LCS.prototype.forEachPositionInSnake = function(left, right, callback, T) {
        var k = right.k;
        var x = (k > left.k) ? left.x + 1 : left.x;
        var n = right.x;

        while (x < n) {
            callback.call(T, x, x-k);
            x++;
        }
    };


    /**
     * Create a new KPoint instance.
     *
     * A KPoint represents a point identified by an x-coordinate and the
     * number of the k-line it is located at.
     *
     * @constructor
     */
    var KPoint = function(x, k) {
        /**
         * The x-coordinate of the k-point.
         */
        this.x = x;

        /**
         * The k-line on which the k-point is located at.
         */
        this.k = k;
    };


    /**
     * Return a new copy of this k-point.
     */
    KPoint.prototype.copy = function() {
        return new KPoint(this.x, this.k);
    };


    /**
     * Set the values of a k-point.
     */
    KPoint.prototype.set = function(x, k) {
        this.x = x;
        this.k = k;
        return this;
    };


    /**
     * Translate this k-point by adding the values of the given k-point.
     */
    KPoint.prototype.translate = function(other) {
        this.x += other.x;
        this.k += other.k;
        return this;
    };


    /**
     * Move the point left by d units
     */
    KPoint.prototype.moveleft = function(d) {
        this.x -= d || 1;
        this.k -= d || 1;
        return this;
    };


    /**
     * Move the point right by d units
     */
    KPoint.prototype.moveright = function(d) {
        this.x += d || 1;
        this.k += d || 1;
        return this;
    };


    /**
     * Move the point up by d units
     */
    KPoint.prototype.moveup = function(d) {
        this.k -= d || 1;
        return this;
    };


    /**
     * Move the point down by d units
     */
    KPoint.prototype.movedown = function(d) {
        this.k += d || 1;
        return this;
    };


    /**
     * Returns true if the given k-point has equal values
     */
    KPoint.prototype.equal = function(other) {
        return (this.x === other.x && this.k === other.k);
    };




    /**
     * Create a new LCS Limit instance. This is a pure data object which holds
     * precalculated parameters for the lcs algorithm.
     *
     * @constructor
     */
    var Limit = function(left, right) {
        this.left = left;
        this.right = right;
        this.delta = right.k - left.k;
        this.N = right.x - left.x;
        this.M = this.N - this.delta;
        this.dmax = this.N + this.M;
    };

    function Anchor(root, base, index) {
        if (!root) {
            throw new Error('Parameter error: need a reference to the tree root');
        }

        if (!base || (root === base && typeof index === 'undefined')) {
            this.base = undefined;
            this.target = root;
            this.index = undefined;
        }
        else if (typeof index === 'undefined') {
            this.base = base.par;
            this.target = base;
            this.index = base.childidx;
        }
        else {
            this.base = base;
            this.target = base.children[index];
            this.index = index;
        }
    }


    /**
     * @constant
     */
    var UPDATE_NODE_TYPE = 1;

    /**
     * @constant
     */
    var UPDATE_FOREST_TYPE = 2;

    /**
     * Private utility class: Creates a new ParameterBuffer instance.
     *
     * @constructor
     */
    function ParameterBuffer(callback, T) {
        this.callback = callback;
        this.T = T;
        this.removes = [];
        this.inserts = [];
    }


    /**
     * Append an item to the end of the buffer
     */
    ParameterBuffer.prototype.pushRemove = function(item) {
        this.removes.push(item);
    };


    /**
     * Append an item to the end of the buffer
     */
    ParameterBuffer.prototype.pushInsert = function(item) {
        this.inserts.push(item);
    };


    /**
     * Invoke callback with the contents of the buffer array and empty the
     * buffer afterwards.
     */
    ParameterBuffer.prototype.flush = function() {
        if (this.removes.length > 0 || this.inserts.length > 0) {
            this.callback.call(this.T, this.removes, this.inserts);
            this.removes = [];
            this.inserts = [];
        }
    };

    /**
     * Utility class to construct a sequence of attached operations from a
     * matching.
     *
     * @constructor
     */
    function DeltaCollector(matching, root_a, root_b) {
        this.matching = matching;
        this.root_a = root_a;
        this.root_b = root_b || matching.get(root_a);
    }


    /**
     * Default equality test. Override this method if you need to test other
     * node properties instead/beside node value.
     */
    DeltaCollector.prototype.equals = function(a, b) {
        return a.value === b.value;
    };


    /**
     * Invoke a callback for each changeset detected between tree a and tree b
     * according to the given matching.
     *
     * @param callback  A function(type, path, removes, inserts) called
     *                  for each detected set of changes.
     * @param T         Context object bound to "this" when the callback is
     * @param root_a    (internal use) Root node in tree a
     * @param root_b    (internal use) Root node in tree b
     *                  invoked.
     * @param path      (internal use) current path relative to base node. Used
     *                  from recursive calls.
     *
     */
    DeltaCollector.prototype.forEachChange = function(callback, T, root_a, root_b,
                                                      path) {
        var parambuf, i, k, a_nodes, b_nodes, a, b, op, me = this;

        // Initialize stuff if not provided
        path = path || [];
        root_a = root_a || this.root_a;
        root_b = root_b || this.root_b;

        if (root_a !== this.matching.get(root_b)) {
            throw new Error('Parameter error, root_a and root_b must be partners');
        }

        // Flag node-update if value of partners do not match
        if (!this.equals(root_a, root_b)) {
            op = new AttachedOperation(
                new Anchor(this.root_a, root_a),
                UPDATE_NODE_TYPE,
                path.slice(),
                [root_a], [root_b]);
            callback.call(T, op);
        }

        // Operation aggregator for subtree changes
        parambuf = new ParameterBuffer(function(removes, inserts) {
            var start = i - removes.length;
            var op = new AttachedOperation(
                new Anchor(me.root_a, root_a, start),
                UPDATE_FOREST_TYPE,
                path.concat(start),
                removes, inserts);
            callback.call(T, op);
        });


        // Descend one level
        a_nodes = root_a.children;
        b_nodes = root_b.children;
        i = 0; k = 0;
        while (a_nodes[i] || b_nodes[k]) {
            a = a_nodes[i];
            b = b_nodes[k];

            if (a && !this.matching.get(a)) {
                parambuf.pushRemove(a);
                i++;
            }
            else if (b && !this.matching.get(b)) {
                parambuf.pushInsert(b);
                k++;
            }
            else if (a && b && a === this.matching.get(b)) {
                // Flush item aggregators
                parambuf.flush();

                // Recurse
                this.forEachChange(callback, T, a, b, path.concat(i));

                i++;
                k++;
            }
            else {
                throw new Error('Matching is not consistent.');
            }
        }

        parambuf.flush();

        return;
    };


    /**
     * Construct a new attached operation instance. An attached operation is always
     * bound to a tree-node identified thru the anchor.
     *
     * @constructor
     */
    function AttachedOperation(anchor, type, path, remove, insert, handler) {
        /**
         * The anchor where the operation is attached
         */
        this.anchor = anchor;


        /**
         * The operation type, one of UPDATE_NODE_TYPE, UPDATE_FOREST_TYPE
         */
        this.type = type;


        /**
         * An array of integers representing the top-down path from the root
         * node to the anchor of this operation. The anchor point always is
         * the first position after the leading context values. For insert
         * operations it will must point to the first element of the tail
         * context.
         */
        this.path = path;


        /**
         * Null (insert), one tree.Node (update) or sequence of nodes (delete)
         */
        this.remove = remove;


        /**
         * Null (remove), one tree.Node (update) or sequence of nodes (insert)
         */
        this.insert = insert;


        /**
         * A handler object used to toggle operation state in the document. I.e.
         * apply and unapply the operation.
         */
        this.handler = handler;
    }

    /**
     * @fileoverview    Implementation of the "skelmatch" tree matching algorithm.
     *
     * This algorithm is heavily inspired by the XCC tree matching algorithm by
     * Sebastian Rönnau and Uwe M. Borghoff. It shares the idea that the
     * interesting bits are found towards the bottom of the tree.
     *
     * Skel-match divides the problem of finding a partial matching between two
     * structured documents represented by ordered trees into two subproblems:
     * 1.   Detect changes in document content (Longest Common Subsequence among
     *      leaf-nodes).
     * 2.   Detect changes in remaining document structure.
     *
     * By default leaf-nodes are considered content, and internal nodes are
     * treated as structure.
     */


    /**
     * Create a new instance of the XCC diff implementation.
     *
     * @param {tree.Node} a Root node of original tree
     * @param {tree.Node} b Root node of changed tree
     *
     * @constructor
     * @name skelmatch.Diff
     */
    function Diff(a, b) {
        this.a = a; // Root node of tree a
        this.b = b; // Root node of tree b
    }


    /**
     * Create a matching between the two nodes using the skelmatch algorithm
     *
     * @param {tree.Matching} matching A tree matching which will be populated by
     *         diffing tree a and b.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.matchTrees = function(matching) {
        // Associate root nodes
        matching.put(this.b, this.a);

        this.matchContent(matching);
        this.matchStructure(matching);
    };


    /**
     * Return true if the given node should be treated as a content node. Override
     * this method in order to implement custom logic to decide whether a node
     * should be examined during the initial LCS (content) or during the second
     * pass. Default: Return true for leaf-nodes.
     *
     * @param {tree.Node} The node which should be examined.
     *
     * @return {boolean} True if the node is a content-node, false otherwise.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.isContent = function(node) {
        return (node.children.length === 0);
    };


    /**
     * Return true if the given node should be treated as a structure node.
     * Default: Return true for internal nodes.
     *
     * @param {tree.Node} The node which should be examined.
     *
     * @return {boolean} True if the node is a content-node, false otherwise.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.isStructure = function(node) {
        return !this.isContent(node);
    };


    /**
     * Default equality test for node values. Override this method if you need to
     * test other node properties instead/beside node value.
     *
     * @param {tree.Node} a Candidate node from tree a
     * @param {tree.Node} b Candidate node from tree b
     *
     * @return {boolean} Return true if the value of the two nodes is equal.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.equals = function(a, b) {
        return (a.value === b.value);
    };


    /**
     * Default equality test for content nodes. Also test all descendants of a and
     * b for equality. Override this method if you want to use tree hashing for
     * this purpose.
     *
     * @param {tree.Node} a Candidate node from tree a
     * @param {tree.Node} b Candidate node from tree b
     *
     * @return {boolean} Return true if the value of the two nodes is equal.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.equalContent = function(a, b) {
        var i;

        if (a.children.length !== b.children.length) {
            return false;
        }
        for (i = 0; i < a.children.length; i++) {
            if (!this.equalContent(a.children[i], b.children[i])) {
                return false;
            }
        }

        return this.equals(a, b);
    };


    /**
     * Default equality test for structure nodes. Return true if ancestors either
     * have the same node value or if they form a pair. Override this method if you
     * want to use tree hashing for this purpose.
     *
     * @param {tree.Node} a Candidate node from tree a
     * @param {tree.Node} b Candidate node from tree b
     *
     * @return {boolean} Return true if the value of the two nodes is equal.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.equalStructure = function(matching, a, b) {
        if (!matching.get(a) && !matching.get(b)) {
            // Return true if all ancestors fullfill the requirement and if the
            // values of a and b are equal.
            return this.equalStructure(matching, a.par, b.par) && this.equals(a, b);
        }
        else {
            // Return true if a and b form a pair.
            return a === matching.get(b);
        }
    };


    /**
     * Return true if a pair is found in the ancestor chain of a and b.
     *
     * @param {tree.Matching} matching A tree matching which will be populated by
     *         diffing tree a and b.
     * @param {tree.Node} a Candidate node from tree a
     * @param {tree.Node} b Candidate node from tree b
     *
     * @return {boolean} Return true if a pair is found in the ancestor chain.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.matchingCheckAncestors = function(matching, a, b) {
        if (!a || !b) {
            return false;
        }
        else if (!matching.get(a) && !matching.get(b)) {
            return this.matchingCheckAncestors(matching, a.par, b.par);
        }
        else {
            return a === matching.get(b);
        }
    };


    /**
     * Put a and b and all their unmatched ancestors into the matching.
     *
     * @param {tree.Matching} matching A tree matching which will be populated by
     *         diffing tree a and b.
     * @param {tree.Node} a Candidate node from tree a
     * @param {tree.Node} b Candidate node from tree b
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.matchingPutAncestors = function(matching, a, b) {
        if (!a || !b) {
            throw new Error('Parameter error: may not match undefined tree nodes');
        }
        else if (!matching.get(a) && !matching.get(b)) {
            this.matchingPutAncestors(matching, a.par, b.par);
            matching.put(a, b);
        }
        else if (a !== matching.get(b)) {
            throw new Error('Parameter error: fundamental matching rule violated.');
        }
    };


    /**
     * Identify unchanged leaves by comparing them using myers longest common
     * subsequence algorithm.
     *
     * @param {tree.Matching} matching A tree matching which will be populated by
     *         diffing tree a and b.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.matchContent = function(matching) {
        var a_content = [],
            b_content = [],
            lcsinst = new LCS(a_content, b_content);

        // Leaves are considered equal if their values match and if they have
        // the same tree depth. Need to wrap the equality-test function into
        // a closure executed immediately in order to maintain correct context
        // (rename 'this' into 'that').
        lcsinst.equals = (function(that){
            return function(a, b) {
                return a.depth === b.depth && that.equalContent(a, b);
            };
        }(this));

        // Populate leave-node arrays.
        this.a.forEachDescendant(function(n) {
            if (this.isContent(n)) a_content.push(n);
        }, this);
        this.b.forEachDescendant(function(n) {
            if (this.isContent(n)) b_content.push(n);
        }, this);

        // Identify structure-preserving changes. Run lcs over leave nodes of
        // tree a and tree b. Associate the identified leaf nodes and also
        // their ancestors except if this would result in structure-affecting
        // change.
        lcsinst.forEachCommonSymbol(function(x, y) {
            var a = a_content[x], b = b_content[y];

            // Verify that ancestor chain allows that a and b to form a pair.
            if (this.matchingCheckAncestors(matching, a, b)) {
                // Record nodes a and b and all of their ancestors in the
                // matching if and only if the nearest matched ancestors are
                // partners.
                this.matchingPutAncestors(matching, a, b);
            }
        }, this);
    };


    /**
     * Return an array of the bottom-most structure-type nodes beneath the given
     * node.
     *
     * @param {tree.Node} node The internal node from where the search should
     *         start.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.collectBones = function(node) {
        var result = [], outer, i = 0;

        if (this.isStructure(node)) {
            for (i = 0; i < node.children.length; i++) {
                outer = this.collectBones(node.children[i]);
                Array.prototype.push.apply(outer);
            }
            if (result.length === 0) {
                // If we do not have any structure-type descendants, this node is
                // the outer most.
                result.push(node);
            }
        }

        return result;
    };


    /**
     * Invoke the given callback with each sequence of unmatched nodes.
     *
     * @param {tree.Matching}   matching  A partial matching
     * @param {Array}           a_sibs    A sequence of siblings from tree a
     * @param {Array}           b_sibs    A sequence of siblings from tree b
     * @param {function}        callback  A function (a_nodes, b_nodes, a_parent, b_parent)
     *         called for every consecutive sequence of nodes from a_sibs and
     *         b_sibs seperated by one or more node pairs.
     * @param {Object}          T         Context object bound to "this" when the
     *         callback is invoked.
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.forEachUnmatchedSequenceOfSiblings = function(matching,
                                                                 a_sibs, b_sibs, callback, T)
    {
        var a_xmatch = [],  // Array of consecutive sequence of unmatched nodes
        // from a_sibs.
            b_xmatch = [],  // Array of consecutive sequence of unmatched nodes
                            // from b_sibs.
            i = 0,      // Array index into a_sibs
            k = 0,      // Array index into b_sibs
            a,          // Current candidate node in a_sibs
            b;          // Current candidate node in b_sibs

        // Loop through a_sibs and b_sibs simultaneously
        while (a_sibs[i] || b_sibs[k]) {
            a = a_sibs[i];
            b = b_sibs[k];

            if (a && !matching.get(a)) {
                // Skip a if above rules did not apply and a is not in the matching
                a_xmatch.push(a);
                i++;
            }
            else if (b && !matching.get(b)) {
                // Skip b if above rules did not apply and b is not in the matching
                b_xmatch.push(b);
                k++;
            }
            else if (a && b && a === matching.get(b)) {
                // Collect nodes at border structure and detect matches
                callback.call(T, a_xmatch, b_xmatch, a, b);
                a_xmatch = [];
                b_xmatch = [];

                // Recurse, both candidates are in the matching
                this.forEachUnmatchedSequenceOfSiblings(matching, a.children, b.children, callback, T);
                i++;
                k++;
            }
            else {
                // Both candidates are in the matching but they are no partners.
                // This is impossible, bail out.
                throw new Error('Matching is not consistent');
            }
        }

        if (a_xmatch.length > 0 || b_xmatch.length > 0) {
            callback.call(T, a_xmatch, b_xmatch, a, b);
        }
    };


    /**
     * Traverse a partial matching and detect equal structure-type nodes between
     * matched content nodes.
     *
     * @param {tree.Matching}   matching  A partial matching
     *
     * @memberOf skelmatch.Diff
     */
    Diff.prototype.matchStructure = function(matching) {
        // Collect unmatched sequences of siblings from tree a and b. Run lcs over
        // bones for each.
        this.forEachUnmatchedSequenceOfSiblings(matching, this.a.children,
            this.b.children, function(a_nodes, b_nodes) {
                var a_bones = [],
                    b_bones = [],
                    lcsinst = new LCS(a_bones, b_bones);

                // Override equality test.
                lcsinst.equals = (function(that){
                    return function(a, b) {
                        return that.equalStructure(matching, a, b);
                    };
                }(this));

                // Populate bone array
                a_nodes.forEach(function(n) {
                    Array.prototype.push.apply(a_bones, this.collectBones(n));
                }, this);
                b_nodes.forEach(function(n) {
                    Array.prototype.push.apply(b_bones, this.collectBones(n));
                }, this);

                // Identify structure-preserving changes. Run lcs over lower bone ends
                // in tree a and tree b. Associate the identified nodes and also their
                // ancestors except if this would result in structure-affecting change.
                lcsinst.forEachCommonSymbol(function(x, y) {
                    var a = a_bones[x], b = b_bones[y];

                    // Verify that ancestor chain allows that a and b to form a pair.
                    if (this.matchingCheckAncestors(matching, a, b)) {
                        // Record nodes a and b and all of their ancestors in the
                        // matching if and only if the nearest matched ancestors are
                        // partners.
                        this.matchingPutAncestors(matching, a, b);
                    }
                }, this);
            }, this);
    };

    window["AscCommon"] = window["AscCommon"] || {};
    window["AscCommon"].Diff = Diff;
    window["AscCommon"].LCS = LCS;
    window["AscCommon"].KPoint = KPoint;
    window["AscCommon"].Limit = Limit;
    window["AscCommon"].Anchor = Anchor;
    window["AscCommon"].ParameterBuffer = ParameterBuffer;
    window["AscCommon"].DeltaCollector = DeltaCollector;
    window["AscCommon"].AttachedOperation = AttachedOperation;
    window["AscCommon"].UPDATE_FOREST_TYPE = UPDATE_FOREST_TYPE;
})(window);
