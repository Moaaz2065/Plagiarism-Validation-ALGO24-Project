using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static Project.Program;

namespace Project
{
    #region PriorityQ Class
    public class PriorityQueue<T> where T : IComparable<T>
    {
        private List<T> heap = new List<T>();

        public int Count => heap.Count;

        public void Enqueue(T item)
        {
            heap.Add(item);
            int i = heap.Count - 1;
            while (i > 0)
            {
                int parent = (i - 1) / 2;
                if (heap[parent].CompareTo(heap[i]) <= 0)
                    break;

                Swap(parent, i);
                i = parent;
            }
        }

        public T Dequeue()
        {
            int lastIndex = heap.Count - 1;
            T frontItem = heap[0];
            heap[0] = heap[lastIndex];
            heap.RemoveAt(lastIndex);

            --lastIndex;
            int parent = 0;
            while (true)
            {
                int leftChild = parent * 2 + 1;
                if (leftChild > lastIndex)
                    break;

                int rightChild = leftChild + 1;
                if (rightChild <= lastIndex && heap[leftChild].CompareTo(heap[rightChild]) > 0)
                    leftChild = rightChild;

                if (heap[parent].CompareTo(heap[leftChild]) <= 0)
                    break;

                Swap(parent, leftChild);
                parent = leftChild;
            }

            return frontItem;
        }

        private void Swap(int i, int j)
        {
            T temp = heap[i];
            heap[i] = heap[j];
            heap[j] = temp;
        }
    }
    #endregion

    class Program
    {
        static string projPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;
        static List<Edge> edges;
        static Dictionary<int, List<int>> adj;
        static Dictionary<int, bool> vis;
        static Dictionary<int, int> sumWeight, mpEdges, GroupRank;
        static Stopwatch MstTime, groupingTime, allTime;
        static string mstAlgorithm = "";

        #region Disjoint-Set Union
        public struct DSU
        {
            public Dictionary<int, int> parent;
            public Dictionary<int, int> GroupSz;

            public DSU()
            {
                parent = new Dictionary<int, int>();
                GroupSz = new Dictionary<int, int>();
            }
            public void Add(int node)
            {
                if (parent.ContainsKey(node))
                    return;
                parent.Add(node, node);
                GroupSz.Add(node, 1);
            }
            public int FindParent(int node)
            {
                if(parent[node] == node)
                    return node;
                return parent[node] = FindParent(parent[node]);
            }

            public bool Union(int a, int b)
            {
                int Pa = FindParent(a), Pb = FindParent(b);
                if (Pa == Pb)
                    return false;

                if (GroupSz[Pa] >= GroupSz[Pb])
                {
                    parent[Pb] = Pa;
                    GroupSz[Pa] += GroupSz[Pb];
                }
                else
                {
                    parent[Pa] = Pb;
                    GroupSz[Pb] += GroupSz[Pa];
                }
                return true;
            }
        }
        #endregion

        #region Edge
        struct Edge : IComparable<Edge>
        {
            public string NodeA, NodeB;
            public int percentageA, percentageB, EdgeWeight, Lines, A, B;
            public object NodeAhyper, NodeBhyper;

            public Edge(string NodeA, string NodeB, int percentageA, int percentageB, int Lines, object hyper1, object hyper2, int A, int B) : this()
            {
                this.NodeA = NodeA;
                this.NodeB = NodeB;
                this.percentageA = percentageA;
                this.percentageB = percentageB;
                this.Lines = Lines;
                this.NodeAhyper = hyper1;
                this.NodeBhyper = hyper2;
                this.A = A; this.B = B;
                EdgeWeight = Math.Max(percentageA, percentageB);
            }
            public int CompareTo(Edge other)
            {
                if (this.EdgeWeight == other.EdgeWeight)
                    return other.Lines.CompareTo(this.Lines);

                return other.EdgeWeight.CompareTo(this.EdgeWeight);
            }
            public int CompareAfterGrouping(Edge other)
            {
                if (GroupRank[this.A] == GroupRank[other.A])
                    return other.Lines.CompareTo(this.Lines);
                

                return GroupRank[this.A].CompareTo(GroupRank[other.A]);
            }
        }
        static Edge createEdge(object Node1, object Node2, object Lines, object hyper1, object hyper2)
        {
            string A = Node1.ToString(), B = Node2.ToString();
            string PA = A.Substring(A.IndexOf('(') + 1, A.IndexOf('%') - A.IndexOf('(') - 1);
            string PB = B.Substring(B.IndexOf('(') + 1, B.IndexOf('%') - B.IndexOf('(') - 1);
            A = A.Substring(0, A.IndexOf('('));
            B = B.Substring(0, B.IndexOf('('));
            string tmp1 = "";
            foreach (var i in A)
                if (char.IsDigit(i))
                    tmp1 += i;
            string tmp2 = "";
            foreach (var i in B)
                if (char.IsDigit(i))
                    tmp2 += i;
            return new Edge(A, B, int.Parse(PA), int.Parse(PB), int.Parse(Lines.ToString()), hyper1, hyper2, int.Parse(tmp1), int.Parse(tmp2));
        }
        #endregion

        #region MST Algorithms
        static List<Edge> kruskalMST()
        {
            DSU dsu = new DSU();
            List<Edge> ret = new List<Edge>();
            edges.Sort();
            foreach (var i in edges)
            {
                dsu.Add(i.A);
                dsu.Add(i.B);
                if (dsu.Union(i.A, i.B))
                    ret.Add(i);
            }
            return ret;
        }
        static List<Edge> primsMST()
        {
            Dictionary<int, List<Tuple<Edge, int>>> adjacency = new Dictionary<int, List<Tuple<Edge, int>>>();
            vis = new Dictionary<int, bool>();
            foreach (var edge in edges)
            {
                if (!adjacency.ContainsKey(edge.A))
                {
                    adjacency.Add(edge.A, new List<Tuple<Edge, int>>());
                    vis.Add(edge.A, false);
                }
                if (!adjacency.ContainsKey(edge.B))
                {
                    adjacency.Add(edge.B, new List<Tuple<Edge, int>>());
                    vis.Add(edge.B, false);
                }
                adjacency[edge.A].Add(Tuple.Create(edge, edge.B));
                adjacency[edge.B].Add(Tuple.Create(edge, edge.A));
            }
            List<Edge> ret = new List<Edge>();
            foreach (var node in adjacency.Keys)
            {
                if (vis[node])
                    continue;
                PriorityQueue<(Edge, int)> pq = new PriorityQueue<(Edge, int)>();
                pq.Enqueue((new Edge(), node));
                while (pq.Count > 0)
                {
                    var p = pq.Dequeue();

                    if (vis[p.Item2])
                        continue;
                    vis[p.Item2] = true;
                    if (p.Item1.Lines != 0)
                        ret.Add(p.Item1);
                    foreach (var child in adjacency[p.Item2])
                    {
                        if (!vis[child.Item2])
                        {
                            pq.Enqueue((child.Item1, child.Item2));
                        }
                    }
                }
            }
            return ret;
        }
        #endregion

        #region Grouping DFS
        struct Group : IComparable<Group>
        {
            public List<int> items;
            public double similarity;

            public Group()
            {
                items = new List<int>();
                similarity = 0.0;
            }
            public int CompareTo(Group other)
            {
                if (this.similarity == other.similarity)
                    return other.items.Count.CompareTo(this.items.Count);
                return other.similarity.CompareTo(this.similarity);
            }
        }
        static void DFS(int node, ref double totalWeight, ref double componentSize, ref Group myG)
        {
            vis[node] = true;
            totalWeight += sumWeight[node];
            componentSize += mpEdges[node];
            myG.items.Add(node);

            foreach (var child in adj[node])
            {
                if (!vis[child])
                    DFS(child, ref totalWeight, ref componentSize, ref myG);
            }
        }

        static List<Group> groupingItems()
        {
            adj = new Dictionary<int, List<int>>();
            vis = new Dictionary<int, bool>();
            sumWeight = new Dictionary<int, int>();
            mpEdges = new Dictionary<int, int>();
            foreach (var edge in edges)
            {
                if (!adj.ContainsKey(edge.A))
                {
                    adj.Add(edge.A, new List<int>());
                    vis.Add(edge.A, false);
                    sumWeight.Add(edge.A, 0);
                    mpEdges.Add(edge.A, 0);
                }
                if (!adj.ContainsKey(edge.B))
                {
                    adj.Add(edge.B, new List<int>());
                    vis.Add(edge.B, false);
                    sumWeight.Add(edge.B, 0);
                    mpEdges.Add(edge.B, 0);
                }
                adj[edge.A].Add(edge.B);
                adj[edge.B].Add(edge.A);

                sumWeight[edge.A] += edge.percentageA;
                sumWeight[edge.B] += edge.percentageB;

                mpEdges[edge.A]++;
            }
            List<Group> groups = new List<Group>();
            foreach (var i in adj.Keys)
            {
                if (!vis[i])
                {
                    double totalWeight = 0.0f;
                    double componentSz = 0.0f;
                    Group g = new Group();

                    DFS(i, ref totalWeight, ref componentSz, ref g);

                    g.similarity = Math.Round(totalWeight / (2.0 * componentSz), 1);

                    groups.Add(g);
                }
            }

            return groups;
        }
        #endregion

        #region Reading/Creating Excel File
        public static void readExcelFile(string fileName)
        {
            edges = new List<Edge>();
            using (var package = new ExcelPackage(new System.IO.FileInfo(projPath + fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value == null || worksheet.Cells[row, 2].Value == null)
                        break;
                    object hyper1 = worksheet.Cells[row, 1].Hyperlink;
                    object hyper2 = worksheet.Cells[row, 2].Hyperlink;
                    edges.Add(createEdge(worksheet.Cells[row, 1].Value, worksheet.Cells[row, 2].Value, worksheet.Cells[row, 3].Value, hyper1, hyper2));
                }
            }
        }
        public static void createExcelFileMST(string fileName)
        {
            MstTime = new Stopwatch();
            MstTime.Start();
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells["A1"].Value = "File 1";
                worksheet.Cells["B1"].Value = "File 2";
                worksheet.Cells["C1"].Value = "Line Matches";

                List<Edge> Answer;
                Answer = (mstAlgorithm == "Kruskal") ? kruskalMST() : primsMST();



                Answer.Sort((x, y) => x.CompareAfterGrouping(y));


                int curr = 2;

                foreach (Edge edge in Answer)
                {
                    worksheet.Cells[$"A{curr}"].Value = edge.NodeA + $" ({edge.percentageA}%)";
                    if (edge.NodeAhyper != null)
                        worksheet.Cells[$"A{curr}"].Hyperlink = new ExcelHyperLink(edge.NodeAhyper.ToString());
                    worksheet.Cells[$"A{curr}"].Style.Font.UnderLine = true;
                    worksheet.Cells[$"A{curr}"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                    worksheet.Cells[$"B{curr}"].Value = edge.NodeB + $" ({edge.percentageB}%)";
                    if (edge.NodeBhyper != null)
                        worksheet.Cells[$"B{curr}"].Hyperlink = new ExcelHyperLink(edge.NodeBhyper.ToString());
                    worksheet.Cells[$"B{curr}"].Style.Font.UnderLine = true;
                    worksheet.Cells[$"B{curr}"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                    worksheet.Cells[$"C{curr}"].Value = edge.Lines;
                    curr++;
                }
                worksheet.Cells.AutoFitColumns();

                FileInfo excelFile = new FileInfo(projPath + @$"\Output\{fileName}.xlsx");
                excelPackage.SaveAs(excelFile);
            }
            MstTime.Stop();
        }
        public static void createExcelFileSTAT(string fileName)
        {
            groupingTime = new Stopwatch();
            groupingTime.Start();
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells["A1"].Value = "Component Index";
                worksheet.Cells["B1"].Value = "Vertices";
                worksheet.Cells["C1"].Value = "Average Similarity";
                worksheet.Cells["D1"].Value = "Component Count";

                List<Group> Answer = groupingItems();
                foreach (Group group in Answer) group.items.Sort();
                Answer.Sort();
                int curr = 2;

                GroupRank = new Dictionary<int, int>();
                foreach (Group group in Answer)
                {
                    worksheet.Cells[$"A{curr}"].Value = curr - 1;
                    worksheet.Cells[$"B{curr}"].Value = "";
                    foreach (var item in group.items)
                    {
                        GroupRank.Add(item, curr - 1);
                        if (item == group.items.Last())
                            break;
                        worksheet.Cells[$"B{curr}"].Value += $"{item}, ";
                    }
                    worksheet.Cells[$"B{curr}"].Value += $"{group.items.Last()}";
                    worksheet.Cells[$"C{curr}"].Value = group.similarity;
                    worksheet.Cells[$"D{curr}"].Value = group.items.Count;
                    curr++;
                }
                worksheet.Cells.AutoFitColumns();

                FileInfo excelFile = new FileInfo(projPath + @$"\Output\{fileName}.xlsx");
                excelPackage.SaveAs(excelFile);
            }
            groupingTime.Stop();
        }
        #endregion

        #region Sample Run
        public static void runSample()
        {
            for (int i = 1; i <= 6; i++)
            {
                Console.Write($"Sample {i} : ");
                readExcelFile(@$"\Test Cases\Sample\{i}-Input.xlsx");
                createExcelFileSTAT(@$"Sample\SampleTest{i}STAT");
                createExcelFileMST(@$"Sample\SampleTest{i}MST");
                Console.WriteLine("Completed");
                Console.WriteLine($"MST Time : {MstTime.ElapsedMilliseconds} ms   Grouping Time : {groupingTime.ElapsedMilliseconds} ms");
            }


        }
        #endregion

        #region Complete Test
        static void runComplete()
        {
            for (int i = 1; i <= 2; i++)
            {
                allTime = new Stopwatch();
                allTime.Start();
                Console.Write($"Easy Test {i} : ");
                readExcelFile(@$"\Test Cases\Complete\Easy\{i}-Input.xlsx");
                createExcelFileSTAT(@$"Complete\Easy\EasyTest{i} STAT");
                createExcelFileMST(@$"Complete\Easy\EasyTest{i} MST With " + mstAlgorithm);
                Console.WriteLine("Completed");
                allTime.Stop();
                Console.WriteLine($"Grouping Time : {groupingTime.ElapsedMilliseconds} ms   MST Time : {MstTime.ElapsedMilliseconds} ms");
                Console.WriteLine($"Overall Time : {allTime.ElapsedMilliseconds}");
                Console.WriteLine("\n#######################\n");
            }


            for (int i = 1; i <= 2; i++)
            {
                allTime = new Stopwatch();
                allTime.Start();
                Console.Write($"Medium Test {i} : ");
                readExcelFile(@$"\Test Cases\Complete\Medium\{i}-Input.xlsx");
                createExcelFileSTAT(@$"Complete\Medium\MediumTest{i} STAT");
                createExcelFileMST(@$"Complete\Medium\MediumTest{i} MST With " + mstAlgorithm);
                Console.WriteLine("Completed");
                Console.WriteLine($"Stat Time : {groupingTime.ElapsedMilliseconds} ms   MST Time : {MstTime.ElapsedMilliseconds} ms");
                Console.WriteLine($"Overall Time : {allTime.ElapsedMilliseconds}");
                Console.WriteLine("\n#######################\n");

            }

            for (int i = 1; i <= 2; i++)
            {
                allTime = new Stopwatch();
                allTime.Start();
                Console.Write($"Hard Test {i} : ");
                readExcelFile(@$"\Test Cases\Complete\Hard\{i}-Input.xlsx");
                createExcelFileSTAT(@$"Complete\Hard\HardTest{i} STAT");
                createExcelFileMST(@$"Complete\Hard\HardTest{i} MST With " + mstAlgorithm);
                Console.WriteLine("Completed");
                Console.WriteLine($"Stat Time : {groupingTime.ElapsedMilliseconds} ms   MST Time : {MstTime.ElapsedMilliseconds} ms");
                Console.WriteLine($"Overall Time : {allTime.ElapsedMilliseconds}");
                Console.WriteLine("\n#######################\n");

            }
        }
        #endregion

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("Press 1 To Use Kruskal Algorithm In MST ");
            Console.WriteLine("Press 2 To Use Prims Algorithm In MST \n");
            
            string input = Console.ReadLine();
            if (input == "2") mstAlgorithm = "prims";
            else mstAlgorithm = "Kruskal";

            //runSample();
            runComplete();
        }
    }
}