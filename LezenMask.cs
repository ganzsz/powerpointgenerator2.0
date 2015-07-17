using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerpointGenerater2
{
    class LezenMask
    {
        /// <summary>
        /// The lezen-type name
        /// </summary>
        private string[] masks;
        /// <summary>
        /// is it lezen or tekst
        /// </summary>
        private string[] type;

        /// <summary>
        /// Constructor with only an array of names
        /// </summary>
        /// <param name="names">The names to use as lezen-type dias</param>
        public LezenMask(params string[] names)
        {
            this.masks = names;
        }

        /// <summary>
        /// Checks if the given argument is in the list
        /// </summary>
        /// <param name="name">The name to test in the list</param>
        /// <returns>true when the name is a lezen-type command, else false</returns>
        public bool Contains(string name)
        {
            return masks.Contains(name);
        }

        /// <summary>
        /// Return the counter with a name
        /// <returns>the counter if found else -1</returns>
        private int getCounter(string name)
        {
            for (int counter = 0; counter < this.masks.Count(); counter++)
            {
                if (this.masks[counter] == name)
                {
                    return counter;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the type of an lezen-like element
        /// </summary>
        /// <param name="name">the name of the element</param>
        /// <returns>The type name, usually lezen or tekst, or NULL if name is not found</returns>
        public string getType(string name)
        {
            int counter = this.getCounter(name);
            if (counter < 0)
                return null;
            else
            {
                return this.type[counter];
            }
        }
    }
}
