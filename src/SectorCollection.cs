/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using System;
using System.Collections;
using System.Collections.Generic;

namespace OpenMcdf
{
    /// <summary>
    /// Action to implement when transaction support - sector
    /// has to be written to the underlying stream (see specs).
    /// </summary>
    public delegate void Ver3SizeLimitReached();

    /// <inheritdoc />
    /// <summary>
    /// Ad-hoc Heap Friendly sector collection to avoid using 
    /// large array that may create some problem to GC collection 
    /// (see http://www.simple-talk.com/dotnet/.net-framework/the-dangers-of-the-large-object-heap/ )
    /// </summary>
    internal class SectorCollection : IList<Sector>
    {
        private const int MAX_SECTOR_V4_COUNT_LOCK_RANGE = 524287; //0x7FFFFF00 for Version 4
        private const int SLICE_SIZE = 4096;

        private readonly List<ArrayList> _largeArraySlices;

        private bool _sizeLimitReached;

        public SectorCollection()
        {
            _largeArraySlices = new List<ArrayList>();
        }

        private void DoCheckSizeLimitReached()
        {
            if (_sizeLimitReached || (Count - 1 <= MAX_SECTOR_V4_COUNT_LOCK_RANGE))
                return;

            OnVer3SizeLimitReached?.Invoke();
            _sizeLimitReached = true;
        }

        public event Ver3SizeLimitReached OnVer3SizeLimitReached;

        #region IList<T> Members

        public int IndexOf(Sector item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, Sector item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        public Sector this[int index]
        {
            get
            {
                var itemIndex = index / SLICE_SIZE;
                var itemOffset = index % SLICE_SIZE;

                if ((index > -1) && (index < Count))
                {
                    return (Sector) _largeArraySlices[itemIndex][itemOffset];
                }
                throw new ArgumentOutOfRangeException("index", index, "Argument out of range");
            }

            set
            {
                var itemIndex = index / SLICE_SIZE;
                var itemOffset = index % SLICE_SIZE;

                if (index > -1 && index < Count)
                {
                    _largeArraySlices[itemIndex][itemOffset] = value;
                }
                else
                    throw new ArgumentOutOfRangeException("index", index, "Argument out of range");
            }
        }

        #endregion

        #region ICollection<T> Members

        public void Add(Sector item)
        {
            DoCheckSizeLimitReached();

            var itemIndex = Count / SLICE_SIZE;

            if (itemIndex < _largeArraySlices.Count)
            {
                _largeArraySlices[itemIndex].Add(item);
                Count++;
            }
            else
            {
                var ar = new ArrayList(SLICE_SIZE) { item };
                _largeArraySlices.Add(ar);
                Count++;
            }
        }

        public void Clear()
        {
            foreach (var slice in _largeArraySlices)
            {
                slice.Clear();
            }

            _largeArraySlices.Clear();

            Count = 0;
        }

        public bool Contains(Sector item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Sector[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count { get; private set; }

        public bool IsReadOnly => false;

        public bool Remove(Sector item)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region IEnumerable<T> Members

        public IEnumerator<Sector> GetEnumerator()
        {
            for (var i = 0; i < _largeArraySlices.Count; i++)
            {
                for (var j = 0; j < _largeArraySlices[i].Count; j++)
                {
                    yield return (Sector) _largeArraySlices[i][j];
                }
            }
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (var arrayList in _largeArraySlices)
            {
                foreach (var obj in arrayList)
                {
                    yield return obj;
                }
            }
        }

        #endregion
    }
}