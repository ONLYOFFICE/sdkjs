/**
 * The MIT License
 *
 * Copyright (c) 2010-2018 Douglas Duhaime http://douglasduhaime.com
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
var Minhash = function(config) {

    // prime is the smallest prime larger than the largest
    // possible hash value (max hash = 32 bit int)
    this.prime = 4294967311;
    this.maxHash = Math.pow(2, 32) - 1;
    this.count = 0;
    this.countLetters = 0;
    // initialize the hash values as the maximum value
    this.inithashvalues = function() {
        for (var i=0; i<this.numPerm; i++) {
            this.hashvalues.push(this.maxHash);
        }
    };

    // initialize the permutation functions for a & b
    // don't reuse any integers when making the functions
    this.initPermutations_ = function(used) {
        var perms = [];
        for (var j=0; j<this.numPerm; j++) {
            var int = this.randInt();
            while (used[int]) int = this.randInt();
            perms.push(int);
            used[int] = true;
        }
        return perms;
    };

    // initialize the permutation functions for a & b
    // don't reuse any integers when making the functions
    this.initPermutations = function() {
        var used = {};
        this.permA = this.initPermutations_(used);
        this.permB = this.initPermutations_(used);
    };

    // the update function updates internal hashvalues given user data
    this.update = function(aCodes) {
        ++this.count;
        for (var i=0; i<this.hashvalues.length; i++) {
            var a = this.permA[i];
            var b = this.permB[i];
            var hash = (a * this.hash(aCodes) + b) % this.prime;
            if (hash < this.hashvalues[i]) {
                this.hashvalues[i] = hash;
            }
        }
    };

    // hash a string to a 32 bit unsigned int
    this.hash = function(aCodes) {
        var hash = 0;
        if (aCodes.length == 0) {
            return hash + this.maxHash;
        }
        for (var i = 0; i < aCodes.length; i++) {
            var char = aCodes[i];
            hash = ((hash<<5)-hash)+char;
            hash = hash & hash; // convert to a 32bit integer
        }
        return hash + this.maxHash;
    };

    // estimate the jaccard similarity to another minhash
    this.jaccard = function(other) {
        if (this.hashvalues.length != other.hashvalues.length) {
            throw new Error('hashvalue counts differ');
        } else if (this.seed != other.seed) {
            throw new Error('seed values differ');
        }
        var shared = 0;
        for (var i=0; i<this.hashvalues.length; i++) {
            shared += this.hashvalues[i] == other.hashvalues[i];
        }
        return shared / this.hashvalues.length;
    };

    // return a random integer >= 0 and <= maxHash
    this.randInt = function() {
        var x = Math.sin(this.seed++) * this.maxHash;
        return Math.floor((x - Math.floor(x)) * this.maxHash);
    };

    // initialize the minhash
    var config = config || {};
    this.numPerm = config.numPerm || 128;
    this.seed = config.seed || 1;
    this.hashvalues = [];
    this.permA = [];
    this.permB = [];
    // share permutation functions across all minhashes
    this.inithashvalues();
    this.initPermutations();
};
