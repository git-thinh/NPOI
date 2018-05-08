/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/
 
 using System;
using System.Text;
using System.Collections;
using System.IO;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NPOI.POIFS.Common;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;

namespace TestCases.POIFS.Properties
{
    /**
     * Class to Test RootProperty functionality
     *
     * @author Marc Johnson
     */
    [TestClass]
    public class TestRootProperty
    {
        private RootProperty _property;
        private byte[] _testblock;

        /**
         * Constructor TestRootProperty
         *
         * @param name
         */

        public TestRootProperty()
        {

        }

        /**
         * Test constructing RootProperty
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestConstructor()
        {
            CreateBasicRootProperty();
            VerifyProperty();
        }

        private void CreateBasicRootProperty()
        {
            _property = new RootProperty();
            _testblock = new byte[128];
            int index = 0;

            for (; index < 0x40; index++)
            {
                _testblock[index] = (byte)0;
            }
            String name = "Root Entry";
            int limit = Math.Min(31, name.Length);

            _testblock[index++] = (byte)(2 * (limit + 1));
            _testblock[index++] = (byte)0;
            _testblock[index++] = (byte)5;
            _testblock[index++] = (byte)1;
            for (; index < 0x50; index++)
            {
                _testblock[index] = (byte)0xff;
            }
            for (; index < 0x74; index++)
            {
                _testblock[index] = (byte)0;
            }
            _testblock[index++] = unchecked((byte)POIFSConstants.END_OF_CHAIN);
            for (; index < 0x78; index++)
            {
                _testblock[index] = (byte)0xff;
            }
            for (; index < 0x80; index++)
            {
                _testblock[index] = (byte)0;
            }
            byte[] name_bytes = System.Text.Encoding.UTF8.GetBytes(name);

            for (index = 0; index < limit; index++)
            {
                _testblock[index * 2] = name_bytes[index];
            }
        }

        private void VerifyProperty()
        {
            MemoryStream stream = new MemoryStream(512);

            _property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(_testblock.Length, output.Length);
            for (int j = 0; j < _testblock.Length; j++)
            {
                Assert.AreEqual(_testblock[j],
                             output[j], "mismatch at offset " + j);
            }
        }

        /**
         * Test SetSize
         */
        [TestMethod]
        public void TestSetSize()
        {
            for (int j = 0; j < 10; j++)
            {
                CreateBasicRootProperty();
                _property.Size = j;
                Assert.AreEqual(j * 64,
                             _property.Size, "trying block count of " + j);
            }
        }

        /**
         * Test Reading constructor
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestReadingConstructor()
        {
            byte[] input =
        {
            ( byte ) 0x52, ( byte ) 0x00, ( byte ) 0x6F, ( byte ) 0x00,
            ( byte ) 0x6F, ( byte ) 0x00, ( byte ) 0x74, ( byte ) 0x00,
            ( byte ) 0x20, ( byte ) 0x00, ( byte ) 0x45, ( byte ) 0x00,
            ( byte ) 0x6E, ( byte ) 0x00, ( byte ) 0x74, ( byte ) 0x00,
            ( byte ) 0x72, ( byte ) 0x00, ( byte ) 0x79, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x16, ( byte ) 0x00, ( byte ) 0x05, ( byte ) 0x01,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x02, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x20, ( byte ) 0x08, ( byte ) 0x02, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xC0, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x46,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0xC0, ( byte ) 0x5C, ( byte ) 0xE8, ( byte ) 0x23,
            ( byte ) 0x9E, ( byte ) 0x6B, ( byte ) 0xC1, ( byte ) 0x01,
            ( byte ) 0xFE, ( byte ) 0xFF, ( byte ) 0xFF, ( byte ) 0xFF,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00,
            ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00, ( byte ) 0x00
        };

            VerifyReadingProperty(0, input, 0, "Root Entry", "{00020820-0000-0000-C000-000000000046}");
        }

        private void VerifyReadingProperty(int index, byte[] input, int offset,
                                           String name, String sClsId)
        {
            RootProperty property = new RootProperty(index, input,
                                                 offset);
            MemoryStream stream = new MemoryStream(128);
            byte[] expected = new byte[128];

            Array.Copy(input, offset, expected, 0, 128);
            property.WriteData(stream);
            byte[] output = stream.ToArray();

            Assert.AreEqual(128, output.Length);
            for (int j = 0; j < 128; j++)
            {
                Assert.AreEqual(expected[j],
                             output[j], "mismatch at offset " + j);
            }
            Assert.AreEqual(index, property.Index);
            Assert.AreEqual(name, property.Name);
            Assert.IsTrue(!property.Children.MoveNext());
            Assert.AreEqual(property.StorageClsid.ToString(), sClsId);
        }

    }
}