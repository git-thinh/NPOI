
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
   
using System;
using System.Collections;
using NPOI.POIFS.FileSystem;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/**
 * Class to Test POIFSReaderRegistry functionality
 *
 * @author Marc Johnson
 */
[TestClass]
public class TestPOIFSReaderRegistry

{
    private POIFSReaderListener[] listeners =
    {
        new Listener(), new Listener(), new Listener(), new Listener()
    };
    private POIFSDocumentPath[]   paths     =
    {
        new POIFSDocumentPath(), new POIFSDocumentPath(new String[]
        {
            "a"
        }), new POIFSDocumentPath(new String[]
        {
            "b"
        }), new POIFSDocumentPath(new String[]
        {
            "c"
        })
    };
    private String[]              names     =
    {
        "a0", "a1", "a2", "a3"
    };

    /**
     * Constructor TestPOIFSReaderRegistry
     *
     * @param name
     */

    public TestPOIFSReaderRegistry()
    {
        
    }

    /**
     * Test empty registry
     */
    [TestMethod]
    public void TestEmptyRegistry()
    {
        POIFSReaderRegistry registry = new POIFSReaderRegistry();

        for (int j = 0; j < paths.Length; j++)
        {
            for (int k = 0; k < names.Length; k++)
            {
                IEnumerator listeners = registry.GetListeners(paths[ j ],
                                                           names[ k ]);

                Assert.IsTrue(!listeners.MoveNext());
            }
        }
    }

    /**
     * Test mixed registration operations
     */
    [TestMethod]
    public void TestMixedRegistrationOperations()
    {
        POIFSReaderRegistry registry = new POIFSReaderRegistry();

        for (int j = 0; j < listeners.Length; j++)
        {
            for (int k = 0; k < paths.Length; k++)
            {
                for (int n = 0; n < names.Length; n++)
                {
                    if ((j != k) && (k != n))
                    {
                        registry.RegisterListener(listeners[ j ], paths[ k ],
                                                  names[ n ]);
                    }
                }
            }
        }
        for (int k = 0; k < paths.Length; k++)
        {
            for (int n = 0; n < names.Length; n++)
            {
                IEnumerator listeners = registry.GetListeners(paths[ k ],
                                                           names[ n ]);

                if (k == n)
                {
                    Assert.IsTrue(!listeners.MoveNext());
                }
                else
                {
                    ArrayList registeredListeners = new ArrayList();

                    while (listeners.MoveNext())
                    {
                        registeredListeners.Add(listeners.Current);
                    }
                    Assert.AreEqual(this.listeners.Length - 1,
                                 registeredListeners.Count);
                    for (int j = 0; j < this.listeners.Length; j++)
                    {
                        if (j == k)
                        {
                            Assert.IsTrue(!registeredListeners
                                .Contains(this.listeners[ j ]));
                        }
                        else
                        {
                            Assert.IsTrue(registeredListeners
                                .Contains(this.listeners[ j ]));
                        }
                    }
                }
            }
        }
        for (int j = 0; j < listeners.Length; j++)
        {
            registry.RegisterListener(listeners[ j ]);
        }
        for (int k = 0; k < paths.Length; k++)
        {
            for (int n = 0; n < names.Length; n++)
            {
                IEnumerator listeners           =
                    registry.GetListeners(paths[ k ], names[ n ]);
                ArrayList registeredListeners = new ArrayList();

                while (listeners.MoveNext())
                {
                    registeredListeners.Add(listeners.Current);
                }
                Assert.AreEqual(this.listeners.Length,
                             registeredListeners.Count);
                for (int j = 0; j < this.listeners.Length; j++)
                {
                    Assert.IsTrue(registeredListeners
                        .Contains(this.listeners[ j ]));
                }
            }
        }
    }
}
