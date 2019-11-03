/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package io.starter.toolkit

import java.util.*

/*
 *  Copyright 2004 The Apache Software Foundation
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

/**
 * A `List` implementation that is optimised for fast insertions and
 * removals at any index in the list.
 *
 *
 * This list implementation utilises a tree structure internally to ensure that
 * all insertions and removals are O(log n). This provides much faster
 * performance than both an `ArrayList` and a `LinkedList`
 * where elements are inserted and removed repeatedly from anywhere in the list.
 *
 *
 * The following relative performance statistics are indicative of this class:
 *
 * <pre>
 * get  add  insert  iterate  remove
 * TreeList       3    5       1       2       1
 * ArrayList      1    1      40       1      40
 * LinkedList  5800    1     350       2     325
</pre> *
 *
 * `ArrayList` is a good general purpose list implementation. It is
 * faster than `TreeList` for most operations except inserting and
 * removing in the middle of the list. `ArrayList` also uses less
 * memory as `TreeList` uses one object per entry.
 *
 *
 * `LinkedList` is rarely a good choice of implementation.
 * `TreeList` is almost always a good replacement for it, although it
 * does use sligtly more memory.
 *
 * @author Joerg Schmuecker
 * @author Stephen Colebourne
 * @version $Revision: 1.2 $ $Date: 2010/01/26 14:53:49 $
 * @since Commons Collections 3.1
 */
internal class XLSOptimizedTreeList : AbstractList<Any> {

    /**
     * The root node in the AVL tree
     */
    private var root: AVLNode? = null

    /**
     * The current size of the list
     */
    private var size: Int = 0

    // -----------

    /**
     * Constructs a new empty list.
     */
    constructor() : super() {}

    /**
     * Constructs a new empty list that copies the specified list.
     *
     * @param coll the collection to copy
     * @throws NullPointerException if the collection is null
     */
    constructor(coll: Collection<*>) : super() {
        addAll(coll)
    }

    // -----------

    /**
     * Gets the element at the specified index.
     *
     * @param index the index to retrieve
     * @return the element at the specified index
     */
    override fun get(index: Int): Any? {
        checkInterval(index, 0, size - 1)
        return root!![index]!!.value
    }

    /**
     * Gets the current size of the list.
     *
     * @return the current size
     */
    override fun size(): Int {
        return size
    }

    /**
     * Gets an iterator over the list.
     *
     * @return an iterator over the list
     */
    override fun iterator(): Iterator<*> {
        // override to go 75% faster
        return listIterator(0)
    }

    /**
     * Gets a ListIterator over the list.
     *
     * @return the new iterator
     */
    override fun listIterator(): ListIterator<*> {
        // override to go 75% faster
        return listIterator(0)
    }

    /**
     * Gets a ListIterator over the list.
     *
     * @param fromIndex the index to start from
     * @return the new iterator
     */
    override fun listIterator(fromIndex: Int): ListIterator<*> {
        // override to go 75% faster
        // cannot use EmptyIterator as iterator.add() must work
        checkInterval(fromIndex, 0, size)
        return TreeListIterator(this, fromIndex)
    }

    /**
     * Searches for the index of an object in the list.
     *
     * @return the index of the object, -1 if not found
     */
    override fun indexOf(`object`: Any?): Int {
        // override to go 75% faster
        return if (root == null) {
            -1
        } else root!!.indexOf(`object`, root!!.relativePosition)
    }

    /**
     * Searches for the presence of an object in the list.
     *
     * @return true if the object is found
     */
    override fun contains(`object`: Any?): Boolean {
        return indexOf(`object`) >= 0
    }

    /**
     * Converts the list into an array.
     *
     * @return the list as an array
     */
    override fun toArray(): Array<Any> {
        // override to go 20% faster
        val array = arrayOfNulls<Any>(size)
        if (root != null) {
            root!!.toArray(array, root!!.relativePosition)
        }
        return array
    }

    // -----------

    /**
     * Adds a new element to the list.
     *
     * @param index the index to add before
     * @param obj   the element to add
     */
    override fun add(index: Int, obj: Any?) {
        modCount++
        checkInterval(index, 0, size)
        if (root == null) {
            root = AVLNode(index, obj, null, null)
        } else {
            root = root!!.insert(index, obj)
        }
        size++
    }

    /**
     * Sets the element at the specified index.
     *
     * @param index the index to set
     * @param obj   the object to store at the specified index
     * @return the previous object at that index
     * @throws IndexOutOfBoundsException if the index is invalid
     */
    override fun set(index: Int, obj: Any?): Any? {
        checkInterval(index, 0, size - 1)
        val node = root!![index]
        val result = node!!.value
        node.value = obj
        return result
    }

    /**
     * Removes the element at the specified index.
     *
     * @param index the index to remove
     * @return the previous object at that index
     */
    override fun remove(index: Int): Any? {
        modCount++
        checkInterval(index, 0, size - 1)
        val result = get(index)
        root = root!!.remove(index)
        size--
        return result
    }

    /**
     * Clears the list, removing all entries.
     */
    override fun clear() {
        modCount++
        root = null
        size = 0
    }

    // -----------

    /**
     * Checks whether the index is valid.
     *
     * @param index      the index to check
     * @param startIndex the first allowed index
     * @param endIndex   the last allowed index
     * @throws IndexOutOfBoundsException if the index is invalid
     */
    private fun checkInterval(index: Int, startIndex: Int, endIndex: Int) {
        if (index < startIndex || index > endIndex) {
            throw IndexOutOfBoundsException("Invalid index:$index, size=$size")
        }
    }

    // -----------

    /**
     * Implements an AVLNode which keeps the offset updated.
     *
     *
     * This node contains the real work. TreeList is just there to implement
     * [java.util.List]. The nodes don't know the index of the object they are
     * holding. They do know however their position relative to their parent node.
     * This allows to calculate the index of a node while traversing the tree.
     *
     *
     * The Faedelung calculation stores a flag for both the left and right child to
     * indicate if they are a child (false) or a link as in linked list (true).
     */
    internal class AVLNode
    /**
     * Constructs a new node with a relative position.
     *
     * @param relativePosition the relative position of the node
     * @param obj              the value for the ndoe
     * @param rightFollower    the node with the value following this one
     * @param leftFollower     the node with the value leading this one
     */
    private constructor(
            /**
             * The relative position, root holds absolute position.
             */
            private var relativePosition: Int,
            /**
             * The stored element.
             */
            /**
             * Gets the value.
             *
             * @return the value of this node
             */
            /**
             * Sets the value.
             *
             * @param obj the value to store
             */
            var value: Any?,
            /**
             * The right child node or the successor if [.rightIsNext].
             */
            private var right: AVLNode?,
            /**
             * The left child node or the predecessor if [.leftIsPrevious].
             */
            private var left: AVLNode?) {
        /**
         * Flag indicating that left reference is not a subtree but the predecessor.
         */
        private var leftIsPrevious: Boolean = false
        /**
         * Flag indicating that right reference is not a subtree but the successor.
         */
        private var rightIsNext: Boolean = false
        /**
         * How many levels of left/right are below this one.
         */
        private var height: Int = 0

        // -----------

        /**
         * Gets the left node, returning null if its a faedelung.
         */
        private val leftSubTree: AVLNode?
            get() = if (leftIsPrevious) null else left

        /**
         * Gets the right node, returning null if its a faedelung.
         */
        private val rightSubTree: AVLNode?
            get() = if (rightIsNext) null else right

        init {
            rightIsNext = true
            leftIsPrevious = true
        }

        /**
         * Locate the element with the given index relative to the offset of the parent
         * of this node.
         */
        operator fun get(index: Int): AVLNode? {
            val indexRelativeToMe = index - relativePosition

            if (indexRelativeToMe == 0) {
                return this
            }

            val nextNode = (if (indexRelativeToMe < 0) leftSubTree else rightSubTree) ?: return null
            return nextNode[indexRelativeToMe]
        }

        /**
         * Locate the index that contains the specified object.
         */
        fun indexOf(`object`: Any?, index: Int): Int {
            if (leftSubTree != null) {
                val result = left!!.indexOf(`object`, index + left!!.relativePosition)
                if (result != -1) {
                    return result
                }
            }
            if (if (value == null) value === `object` else value == `object`) {
                return index
            }
            return if (rightSubTree != null) {
                right!!.indexOf(`object`, index + right!!.relativePosition)
            } else -1
        }

        /**
         * Stores the node and its children into the array specified.
         *
         * @param array the array to be filled
         * @param index the index of this node
         */
        fun toArray(array: Array<Any>, index: Int) {
            array[index] = value
            if (leftSubTree != null) {
                left!!.toArray(array, index + left!!.relativePosition)
            }
            if (rightSubTree != null) {
                right!!.toArray(array, index + right!!.relativePosition)
            }
        }

        /**
         * Gets the next node in the list after this one.
         *
         * @return the next node
         */
        operator fun next(): AVLNode? {
            return if (rightIsNext || right == null) {
                right
            } else right!!.min()
        }

        /**
         * Gets the node in the list before this one.
         *
         * @return the previous node
         */
        fun previous(): AVLNode? {
            return if (leftIsPrevious || left == null) {
                left
            } else left!!.max()
        }

        /**
         * Inserts a node at the position index.
         *
         * @param index is the index of the position relative to the position of the
         * parent node.
         * @param obj   is the object to be stored in the position.
         */
        fun insert(index: Int, obj: Any?): AVLNode {
            val indexRelativeToMe = index - relativePosition

            return if (indexRelativeToMe <= 0) {
                insertOnLeft(indexRelativeToMe, obj)
            } else {
                insertOnRight(indexRelativeToMe, obj)
            }
        }

        private fun insertOnLeft(indexRelativeToMe: Int, obj: Any?): AVLNode {
            var ret = this

            if (leftSubTree == null) {
                setLeft(AVLNode(-1, obj, this, left), null)
            } else {
                setLeft(left!!.insert(indexRelativeToMe, obj), null)
            }

            if (relativePosition >= 0) {
                relativePosition++
            }
            ret = balance()
            recalcHeight()
            return ret
        }

        private fun insertOnRight(indexRelativeToMe: Int, obj: Any?): AVLNode {
            var ret = this

            if (rightSubTree == null) {
                setRight(AVLNode(+1, obj, right, this), null)
            } else {
                setRight(right!!.insert(indexRelativeToMe, obj), null)
            }
            if (relativePosition < 0) {
                relativePosition--
            }
            ret = balance()
            recalcHeight()
            return ret
        }

        /**
         * Gets the rightmost child of this node.
         *
         * @return the rightmost child (greatest index)
         */
        private fun max(): AVLNode {
            return if (rightSubTree == null) this else right!!.max()
        }

        /**
         * Gets the leftmost child of this node.
         *
         * @return the leftmost child (smallest index)
         */
        private fun min(): AVLNode {
            return if (leftSubTree == null) this else left!!.min()
        }

        /**
         * Removes the node at a given position.
         *
         * @param index is the index of the element to be removed relative to the position
         * of the parent node of the current node.
         */
        fun remove(index: Int): AVLNode? {
            val indexRelativeToMe = index - relativePosition

            if (indexRelativeToMe == 0) {
                return removeSelf()
            }
            if (indexRelativeToMe > 0) {
                setRight(right!!.remove(indexRelativeToMe), right!!.right)
                if (relativePosition < 0) {
                    relativePosition++
                }
            } else {
                setLeft(left!!.remove(indexRelativeToMe), left!!.left)
                if (relativePosition > 0) {
                    relativePosition--
                }
            }
            recalcHeight()
            return balance()
        }

        private fun removeMax(): AVLNode? {
            if (rightSubTree == null) {
                return removeSelf()
            }
            setRight(right!!.removeMax(), right!!.right)
            if (relativePosition < 0) {
                relativePosition++
            }
            recalcHeight()
            return balance()
        }

        private fun removeMin(): AVLNode? {
            if (leftSubTree == null) {
                return removeSelf()
            }
            setLeft(left!!.removeMin(), left!!.left)
            if (relativePosition > 0) {
                relativePosition--
            }
            recalcHeight()
            return balance()
        }

        private fun removeSelf(): AVLNode? {
            if (rightSubTree == null && leftSubTree == null)
                return null
            if (rightSubTree == null) {
                if (relativePosition > 0) {
                    left!!.relativePosition += relativePosition + if (relativePosition > 0) 0 else 1
                }
                left!!.max().setRight(null, right)
                return left
            }
            if (leftSubTree == null) {
                right!!.relativePosition += relativePosition - if (relativePosition < 0) 0 else 1
                right!!.min().setLeft(null, left)
                return right
            }

            if (heightRightMinusLeft() > 0) {
                val rightMin = right!!.min()
                value = rightMin.value
                if (leftIsPrevious) {
                    left = rightMin.left
                }
                right = right!!.removeMin()
                if (relativePosition < 0) {
                    relativePosition++
                }
            } else {
                val leftMax = left!!.max()
                value = leftMax.value
                if (rightIsNext) {
                    right = leftMax.right
                }
                left = left!!.removeMax()
                if (relativePosition > 0) {
                    relativePosition--
                }
            }
            recalcHeight()
            return this
        }

        // -----------

        /**
         * Balances according to the AVL algorithm.
         */
        private fun balance(): AVLNode {
            when (heightRightMinusLeft()) {
                1, 0, -1 -> return this
                -2 -> {
                    if (left!!.heightRightMinusLeft() > 0) {
                        setLeft(left!!.rotateLeft(), null)
                    }
                    return rotateRight()
                }
                2 -> {
                    if (right!!.heightRightMinusLeft() < 0) {
                        setRight(right!!.rotateRight(), null)
                    }
                    return rotateLeft()
                }
                else -> throw RuntimeException("tree inconsistent!")
            }
        }

        /**
         * Gets the relative position.
         */
        private fun getOffset(node: AVLNode?): Int {
            return node?.relativePosition ?: 0
        }

        /**
         * Sets the relative position.
         */
        private fun setOffset(node: AVLNode?, newOffest: Int): Int {
            if (node == null) {
                return 0
            }
            val oldOffset = getOffset(node)
            node.relativePosition = newOffest
            return oldOffset
        }

        /**
         * Sets the height by calculation.
         */
        private fun recalcHeight() {
            height = Math.max(if (leftSubTree == null) -1 else leftSubTree!!.height,
                    if (rightSubTree == null) -1 else rightSubTree!!.height) + 1
        }

        /**
         * Returns the height of the node or -1 if the node is null.
         */
        private fun getHeight(node: AVLNode?): Int {
            return node?.height ?: -1
        }

        /**
         * Returns the height difference right - left
         */
        private fun heightRightMinusLeft(): Int {
            return getHeight(rightSubTree) - getHeight(leftSubTree)
        }

        private fun rotateLeft(): AVLNode {
            val newTop = right // can't be faedelung!
            val movedNode = rightSubTree!!.leftSubTree

            val newTopPosition = relativePosition + getOffset(newTop)
            val myNewPosition = -newTop!!.relativePosition
            val movedPosition = getOffset(newTop) + getOffset(movedNode)

            setRight(movedNode, newTop)
            newTop.setLeft(this, null)

            setOffset(newTop, newTopPosition)
            setOffset(this, myNewPosition)
            setOffset(movedNode, movedPosition)
            return newTop
        }

        private fun rotateRight(): AVLNode {
            val newTop = left // can't be faedelung
            val movedNode = leftSubTree!!.rightSubTree

            val newTopPosition = relativePosition + getOffset(newTop)
            val myNewPosition = -newTop!!.relativePosition
            val movedPosition = getOffset(newTop) + getOffset(movedNode)

            setLeft(movedNode, newTop)
            newTop.setRight(this, null)

            setOffset(newTop, newTopPosition)
            setOffset(this, myNewPosition)
            setOffset(movedNode, movedPosition)
            return newTop
        }

        private fun setLeft(node: AVLNode?, previous: AVLNode?) {
            leftIsPrevious = node == null
            left = if (leftIsPrevious) previous else node
            recalcHeight()
        }

        private fun setRight(node: AVLNode?, next: AVLNode?) {
            rightIsNext = node == null
            right = if (rightIsNext) next else node
            recalcHeight()
        }

        // private void checkFaedelung() {
        // AVLNode maxNode = left.max();
        // if (!maxNode.rightIsFaedelung || maxNode.right != this) {
        // throw new RuntimeException(maxNode + " should right-faedel to " + this);
        // }
        // AVLNode minNode = right.min();
        // if (!minNode.leftIsFaedelung || minNode.left != this) {
        // throw new RuntimeException(maxNode + " should left-faedel to " + this);
        // }
        // }
        //
        // private int checkTreeDepth() {
        // int hright = (getRightSubTree() == null ? -1 :
        // getRightSubTree().checkTreeDepth());
        // // Logger.logInfo("checkTreeDepth");
        // // Logger.logInfo(this);
        // // Logger.logInfo(" left: ");
        // // Logger.logInfo(_left);
        // // Logger.logInfo(" right: ");
        // // Logger.logInfo(_right);
        //
        // int hleft = (left == null ? -1 : left.checkTreeDepth());
        // if (height != Math.max(hright, hleft) + 1) {
        // throw new RuntimeException(
        // "height should be max" + hleft + "," + hright + " but is " + height);
        // }
        // return height;
        // }
        //
        // private int checkLeftSubNode() {
        // if (getLeftSubTree() == null) {
        // return 0;
        // }
        // int count = 1 + left.checkRightSubNode();
        // if (left.relativePosition != -count) {
        // throw new RuntimeException();
        // }
        // return count + left.checkLeftSubNode();
        // }
        //
        // private int checkRightSubNode() {
        // AVLNode right = getRightSubTree();
        // if (right == null) {
        // return 0;
        // }
        // int count = 1;
        // count += right.checkLeftSubNode();
        // if (right.relativePosition != count) {
        // throw new RuntimeException();
        // }
        // return count + right.checkRightSubNode();
        // }

        /**
         * Used for debugging.
         */
        override fun toString(): String {
            return ("AVLNode(" + relativePosition + "," + (left != null) + "," + value + ","
                    + (rightSubTree != null) + ", faedelung " + rightIsNext + " )")
        }
    }

    /**
     * A list iterator over the linked list.
     */
    internal class TreeListIterator
    /**
     * Create a ListIterator for a list.
     *
     * @param parent    the parent list
     * @param fromIndex the index to start at
     */
    @Throws(IndexOutOfBoundsException::class)
    constructor(
            /**
             * The parent list
             */
            protected val parent: XLSOptimizedTreeList,
            /**
             * The index of [.next].
             */
            protected var nextIndex: Int) : ListIterator<Any> {
        /**
         * The node that will be returned by [.next]. If this is equal to
         * [AbstractLinkedList.header] then there are no more values to return.
         */
        protected var next: AVLNode? = null
        /**
         * The last node that was returned by [.next] or [.previous].
         * Set to `null` if [.next] or [.previous] haven't
         * been called, or if the node has been removed with [.remove] or a new
         * node added with [.add]. Should be accessed through
         * [.getLastNodeReturned] to enforce this behaviour.
         */
        protected var current: AVLNode? = null
        /**
         * The index of [.current].
         */
        protected var currentIndex: Int = 0
        /**
         * The modification count that the list is expected to have. If the list doesn't
         * have this count, then a [java.util.ConcurrentModificationException] may
         * be thrown by the operations.
         */
        protected var expectedModCount: Int = 0

        init {
            this.expectedModCount = parent.modCount
            this.next = if (parent.root == null) null else parent.root!![nextIndex]
        }

        /**
         * Checks the modification count of the list is the value that this object
         * expects.
         *
         * @throws ConcurrentModificationException If the list's modification count isn't the value that was
         * expected.
         */
        protected fun checkModCount() {
            if (parent.modCount != expectedModCount) {
                // How about Allow it???
                // throw new ConcurrentModificationException();
            }
        }

        override fun hasNext(): Boolean {
            return nextIndex < parent.size
        }

        override fun next(): Any? {
            checkModCount()
            if (!hasNext()) {
                throw NoSuchElementException("No element at index $nextIndex.")
            }
            if (next == null) {
                next = parent.root!![nextIndex]
            }
            val value = next!!.value
            current = next
            currentIndex = nextIndex++
            next = next!!.next()
            return value
        }

        override fun hasPrevious(): Boolean {
            return nextIndex > 0
        }

        override fun previous(): Any? {
            checkModCount()
            if (!hasPrevious()) {
                throw NoSuchElementException("Already at start of list.")
            }
            if (next == null) {
                next = parent.root!![nextIndex - 1]
            } else {
                next = next!!.previous()
            }
            val value = next!!.value
            current = next
            currentIndex = --nextIndex
            return value
        }

        override fun nextIndex(): Int {
            return nextIndex
        }

        override fun previousIndex(): Int {
            return nextIndex() - 1
        }

        override fun remove() {
            checkModCount()
            if (current == null) {
                throw IllegalStateException()
            }
            parent.removeAt(currentIndex)
            current = null
            currentIndex = -1
            nextIndex--
            expectedModCount++
        }

        override fun set(obj: Any) {
            checkModCount()
            if (current == null) {
                throw IllegalStateException()
            }
            current!!.value = obj
        }

        override fun add(obj: Any) {
            checkModCount()
            parent.add(nextIndex, obj)
            current = null
            currentIndex = -1
            nextIndex++
            expectedModCount++
        }
    }

}