Attribute VB_Name = "mp_extsimp1"
'
' mp_extsimp
' Generalization of complex junctions and two ways roads
' from OpenStreetMap data
' Also a toolbox for other generalization procedures
'
' Copyright © 2012-2013 OverQuantum
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Author contacts:
' http://overquantum.livejournal.com
' https://github.com/OverQuantum
'
' Project homepage:
' https://github.com/OverQuantum/mp_extsimp
'
'
' OpenStreetMap data licensed under the Open Data Commons Open Database License (ODbL).
' OpenStreetMap wiki licensed under the Creative Commons Attribution-ShareAlike 2.0 license (CC-BY-SA).
' Please refer to http://www.openstreetmap.org/copyright for details
'
' osm2mp licensed under GPL v2, please refer to http://code.google.com/p/osm2mp/
'
' mp file format (polish format) description from http://www.cgpsmapper.com/manual.htm
'
'
'
'
'history
'2012.10.08 - "challenge accepted", project started
'2012.10.10 - added combining of edges
'2012.10.11 - added checking of type and oneway on combining of edges
'2012.10.12 - added collapsing of junctions
'2012.10.15 - added comments to code, added handling of speedclass
'2012.10.16 - added distance in metres by WGS 84
'2012.10.16 - added joining directions (only duplicate edges and close edges forming V-form)
'2012.10.17 - chg CollapseJunctions now iterative
'2012.10.17 - added inserting near nodes to edges on joining direction (no moving of node)
'2012.10.18 - adding JoinDirections2, closest edge founds, GoByTwoWays started
'2012.10.18 - adding JoinDirections2, joining works, but some mess created...
'2012.10.21 - finished JoinDirections2, handling of circles and deleted void edges
'2012.10.22 - adding CollapseJunctions2, done main part, remain marking edges for cases of crossing if border-nodes >= 2 in the end
'2012.10.22 - adding CollapseJunctions2, marking edges for cases of crossing if border-nodes >= 2 in the end, marking long oneways before
'2012.10.23 - added limiting distance of ShrinkBorderNodes
'2012.10.23 - added check for forward/backward coverage at ends of chains in JoinDirections2
'2012.10.24 - added CheckShortLoop, does not help
'2012.10.29 - added CheckShortLoop, forward/backward coverage in JoinDirections2 for cycles
'2012.10.29 - added DouglasPeucker_total_split
'2012.10.30 - added check lens in CosAngleBetweenEdges, skipping of RouteParamExt, LoopLimit in CJ2
'2012.10.31 - fix forw/back check in JD2 (split to two cycles)
'2012.11.01 - added JoinAcute (from JD), ProjectNode, CompareRoadtype. Looks like RC1 of algo
'2012.11.01 - added Save_MP_2 (w/o rev-dir)
'2012.11.06 - added JD3 - with cluster search
'2012.11.07 - fix loop in D-P, optimized aiming and del in CJ2, added/modified status writing in form1.caption
'2012.11.09 - fix saving2 on oneway=2
'2012.11.12 - added keep of main road label
'2012.11.13 - added TWback/TWforw and checking them in JD3 (unfinished)
'2012.11.14 - added correct naming and speedclass in JD3
 'RC3
'2012.11.14 - hardcoded limits moved to func/sub parameters
'2012.11.14 - unused functions commented
'2012.11.15 - root code moved to module, removed comdlg32.bas
'2012.11.19 - added keeping header of source file, removed writing "; roadtype="
'2012.11.15-20 - adding explaining comments, small fixes
'2012.11.20 - added license and references
'2013.01.08 - added CollapseShortEdges (fix "too close nodes")
'2013.01.28 - fixed deadlock in SaveChain in case of isolated road cycle
'2013.02.28 - added MaxLinkLen to Load_MP()
'2013.03.02 - fixed bug in output degrees<0.1, added LATLON_FORMAT
'2013.03.03 - fixed bug in removing _link on exceeding MaxLinkLen

'2013-09-30 - fixed bug in GetNodeInBboxByCluster on bbox were out of clustered bbox
'2013-09-30 - added consts FORCEWAYSPEED, TRUNK_TYPE, TRUNK_LINK_TYPE and LOAD_NOROUTING
'2013-10-01 - fixed deadlock in DouglasPeucker_chain and _split on isolated two-road cycle
'2013-10-01 - changed form1 to use Form_Load, now program can be run minimized (ex. using "start /MIN")
'2013-10-01 - added JoinCloseNodes(), CombineDuplicateEdgesAll() and RemoveOneWay()
'2013-10-02 - changed edge() to dynamic array
'2013-10-02 - changed realloc of Nodes and Edges to +1M after 1M
'2013-10-02 - added loading of mp Type field, controlled by const LOAD_TYPE
'2013-10-02 - added loading of polygons
'2013-10-02 - fixed ExpandBbox for beyond +-89 degrees
'2013-10-03 - fixed loading of mp Type if no comments
'2013-10-05 - added loading and saving of lined OSM file
'2013-10-05 - added AddNodeToClusterIndex and sorting of nodes by ele
'2013-10-08 - changed control consts to Control_* variables
'2013-10-08 - added showing function params in form caption
'2013-10-08 - added Control_TrunkLinkType
'2013-10-08 - added TrimByBbox, FileLen_safe
'2013-10-09 - added StitchNodes
'2013-10-10 - commented and regression checked with 2013-03-03 version


'TODO:
'*? dump problems of OSM data (1: too long links (ready), 2: ?)
'? 180/-180 safety (currently works fine with planet wide data, but could fail on parallel-wide

Option Explicit

'Conversion degrees to radians and back
Public Const DEGTORAD = 1.74532925199433E-02
Public Const RADTODEG = 57.2957795130823

'WGS84 datum
Public Const DATUM_R_EQUAT = 6378137
Public Const DATUM_R_POLAR = 6356752.3142
Public Const DATUM_R_OVER = 6380000  'for expanding bbox

'OSM highway main types
Public Const HIGHWAY_MOTORWAY = 0
Public Const HIGHWAY_MOTORWAY_LINK = 1
Public Const HIGHWAY_TRUNK = 2
Public Const HIGHWAY_TRUNK_LINK = 3
Public Const HIGHWAY_PRIMARY = 4
Public Const HIGHWAY_PRIMARY_LINK = 5
Public Const HIGHWAY_SECONDARY = 6
Public Const HIGHWAY_SECONDARY_LINK = 7
Public Const HIGHWAY_TERTIARY = 8
Public Const HIGHWAY_TERTIARY_LINK = 9

'OSM highway minor types (should not occur)
Public Const HIGHWAY_LIVING_STREET = 10
Public Const HIGHWAY_RESIDENTIAL = 12
Public Const HIGHWAY_UNCLASSIFIED = 14
Public Const HIGHWAY_SERVICE = 16
Public Const HIGHWAY_TRACK = 18
Public Const HIGHWAY_OTHER = 20
Public Const HIGHWAY_UNKNOWN = 22
Public Const HIGHWAY_UNSPECIFIED = 24

'Masks
Public Const HIGHWAY_MASK_LINK = 1  'all links
Public Const HIGHWAY_MASK_MAIN = 254 'get main type (removes _link)

'Specific marking of edges/nodes for algorithms
Public Const MARK_JUNCTION = 1
Public Const MARK_COLLAPSING = 2
Public Const MARK_AIMING = 4
Public Const MARK_DISTCHECK = 8
Public Const MARK_SIDE1CHECK = 8
Public Const MARK_SIDE2CHECK = 16
Public Const MARK_SIDESCHECK = 24
Public Const MARK_WAVEPASSED = 16
Public Const MARK_NODE_BORDER = -2
Public Const MARK_NODE_OF_JUNCTION = -3
Public Const MARK_NODEID_DELETED = -2

'Format for output coordinates
Public Const LATLON_FORMAT = "0.000000##"

'OSM node - point on Earth with lat/lon coordinates
Public Type node
    lat As Double 'Latitude
    lon As Double 'Longitude
    NodeID As Long 'NodeID from source .mp, -1 - not set, -2 - node killed
    edge() As Long  'all edges (values - indexes in Edges array)
    Edges As Integer 'number of edges, -1 means "not counted"
    EdgesAlloc As Integer 'allocated number of edges
    mark As Long 'internal marker for all-network algo-s
    temp_dist As Double 'internal value for for all-network algo-s
End Type

'Edge - part of OSM way between two nodes
Public Type edge
    node1 As Long 'first node (index in Nodes array)
    node2 As Long 'second node
    roadtype As Byte 'roadtype, see HIGHWAY_ consts
    oneway As Byte '0 - no, 1 - yes ( goes from node1 to node2 )
    mark As Integer 'internal marker for all-network algo-s
    speed As Byte 'speed class (in .mp terms)
    label As String 'label of road (only ref= values currently, not name= )
End Type
    
'Aiming edge - edge for calc centroid of junction
Public Type aimedge
    lat1 As Double
    lon1 As Double
    lat2 As Double
    lon2 As Double
    a As Double 'matrix equation elements
    b As Double
    c As Double
    d As Double
End Type

'bbox, lat/lon min/max rectangle (not 180/-180 safe)
Public Type bbox
    lat_min As Double
    lat_max As Double
    lon_min As Double
    lon_max As Double
End Type

'Element of label statistics "histogramm"
Public Type LabelStat
    Text As String
    count As Long
End Type

Public MPheader As String

'All nodes
Public Nodes() As node
Public NodesAlloc As Long
Public NodesNum As Long

'All road edges
Public Edges() As edge
Public EdgesAlloc As Long
Public EdgesNum As Long

'Aim edges
Public AimEdges() As aimedge
Public AimEdgesAlloc As Long
Public AimEdgesNum As Long

'Array for building chain of indexes
Public Chain() As Long
Public ChainAlloc As Long
Public ChainNum As Long

Public NodeIDMax As Long 'max found NodeID

Public DistanceToSegment_last_case As Long 'case of calc distance during last call of DistanceToSegment()

Public GoByChain_lastedge As Long  'edge, on which GoByChain() function have just passed from node to node

Public SpeedHistogram(10) As Long 'histogramm of speed classes

'Cluster index
Public ClustersLat0 As Double 'min lat-border of clusters
Public ClustersLon0 As Double 'min lon-border of clusters
Public ClustersFirst() As Long 'index of first node of cluster (X*Y)
Public ClustersChain() As Long 'chain of nodes (NodesNum)
Public ClustersLatNum As Long 'num of cluster by lat = X
Public ClustersLonNum As Long 'num of cluster by lon = Y
Public ClustersIndexedNodes As Long 'num of indexed nodes (for continuing BuildNodeClusterIndex)
Public ClustersLast() As Long 'index of last node of cluster - for building index (X*Y)
'for continuing GetNodeInBboxByCluster
Public ClustersFindLastBbox As bbox  'last bbox
Public ClustersFindLastCluster As Long  'last index of cluster
Public ClustersFindLastNode As Long '

'Label statistics, for estimate labels of joined roads
Public LabelStats() As LabelStat
Public LabelStatsNum As Long
Public LabelStatsAlloc As Long

Public EstimateChain_speed As Long 'speed of chain after last call of EstimateChain()
Public EstimateChain_label As String 'label of chain after last call of EstimateChain()

'Indexes of forward and backward ways during two ways joining
Public TWforw() As Long
Public TWback() As Long
Public TWalloc As Long
Public TWforwNum As Long
Public TWbackNum As Long

'Control variables to adjust algorithms, works more like #ifdef
Public Control_ClusterSize As Double 'Cluster size in degrees for ClusterIndex build and search. Dont change without rebuilding index
Public Control_ForceWaySpeed As Long 'Force speed class for ways
Public Control_TrunkType As Long     'MP type for saving highway=trunk
Public Control_TrunkLinkType As Long 'MP type for saving highway=trunk_link
Public Control_PrimaryType As Long   'MP type for saving highway=primary
Public Control_LoadNoRoute As Long   'Load mp object without Route data
Public Control_LoadMPType As Long    'Parse MP type during loading

'Init - init all arrays
Public Sub init()
    
    Control_ClusterSize = 0.05   '0.05 degrees for local maps, 1 for planet-s
    Control_ForceWaySpeed = -1   'set -1 to not force, 0 or more to forcing this value
    Control_TrunkType = 1        'set 1 to be have same as motorway
    Control_PrimaryType = 2      'set 2 to use 0x02 Principal highway
    Control_TrunkLinkType = 9    'set 9 to have same as motorway
    Control_LoadNoRoute = 0      'set 0 to skip no-routing polylines, 1 to load
    Control_LoadMPType = 0       'set 0 to skip mp Type= field, 1 to parse
    
    NodesAlloc = 1000
    ReDim Nodes(NodesAlloc)
    NodesNum = 0
    
    Nodes(0).EdgesAlloc = 3
    ReDim Nodes(0).edge(Nodes(0).EdgesAlloc)
    
    EdgesAlloc = 1000
    ReDim Edges(EdgesAlloc)
    EdgesNum = 0

    ChainAlloc = 10000
    ReDim Chain(ChainAlloc)
    ChainNum = 0
    
    AimEdgesAlloc = 50
    ReDim AimEdges(AimEdgesAlloc)
    AimEdgesNum = 0
    
    LabelStatsAlloc = 30
    ReDim LabelStats(LabelStatsAlloc)
    LabelStatsNum = 0
    
    TWalloc = 100
    ReDim TWforw(TWalloc)
    ReDim TWback(TWalloc)
    TWbackNum = 0
    TWforwNum = 0
End Sub

'Add one node to dynamic array
'Assumed, Nodes(NodesNum) filled with required data prior to call
Public Sub AddNode()
    If NodesNum >= NodesAlloc Then
        'realloc if needed
        If NodesAlloc >= 1000000 Then
            NodesAlloc = NodesAlloc + 1000000
        Else
            NodesAlloc = NodesAlloc * 2
        End If
        ReDim Preserve Nodes(NodesAlloc)
    End If
    NodesNum = NodesNum + 1
    
    Nodes(NodesNum).EdgesAlloc = 3
    ReDim Nodes(NodesNum).edge(Nodes(NodesNum).EdgesAlloc)
End Sub

'Add one edge to dynamic array
'Assumed, Edges(EdgesNum) filled with required data prior to call
Public Sub AddEdge()
    If EdgesNum >= EdgesAlloc Then
        'realloc if needed
        If EdgesAlloc >= 1000000 Then
            EdgesAlloc = EdgesAlloc + 1000000
        Else
            EdgesAlloc = EdgesAlloc * 2
        End If
        ReDim Preserve Edges(EdgesAlloc)
    End If
    EdgesNum = EdgesNum + 1
End Sub

'Add one AimEdge to dynamic array
'Assumed, AimEdges(AimEdgesNum) filled with required data prior to call
Public Sub AddAimEdge()
    If AimEdgesNum >= AimEdgesAlloc Then
        'realloc if needed
        AimEdgesAlloc = AimEdgesAlloc * 2
        ReDim Preserve AimEdges(AimEdgesAlloc)
    End If
    AimEdgesNum = AimEdgesNum + 1
End Sub

'Add one index into chain
Public Sub AddChain(i As Long)
    If ChainNum >= ChainAlloc Then
        'realloc if needed
        ChainAlloc = ChainAlloc * 2
        ReDim Preserve Chain(ChainAlloc)
    End If
    Chain(ChainNum) = i
    ChainNum = ChainNum + 1
End Sub


Public Sub AddEdgeToNode(node1 As Long, edge1 As Long)
    Dim k As Long
    k = Nodes(node1).Edges
    Nodes(node1).edge(k) = edge1
    If Nodes(node1).Edges >= Nodes(node1).EdgesAlloc Then
        'realloc if needed
        Nodes(node1).EdgesAlloc = Nodes(node1).EdgesAlloc + 3
        ReDim Preserve Nodes(node1).edge(Nodes(node1).EdgesAlloc)
    End If
    Nodes(node1).Edges = Nodes(node1).Edges + 1
End Sub


'Join two nodes by new edge
'node1 - start node, 'node2 - end node
'return: index of new edge
Public Function JoinByEdge(node1 As Long, node2 As Long) As Long
    Dim k As Long
    Edges(EdgesNum).node1 = node1
    Edges(EdgesNum).node2 = node2
    
    'add edge to both nodes
    Call AddEdgeToNode(node1, EdgesNum)
    Call AddEdgeToNode(node2, EdgesNum)
    'k = Nodes(node1).Edges
    'Nodes(node1).edge(k) = EdgesNum
    'Nodes(node1).Edges = Nodes(node1).Edges + 1
    'k = Nodes(node2).Edges
    'Nodes(node2).edge(k) = EdgesNum
    'Nodes(node2).Edges = Nodes(node2).Edges + 1
    JoinByEdge = EdgesNum
    Call AddEdge
End Function


'Merge loaded nodes from diffrent ways by NodeID
Public Sub JoinNodesByID()
    Dim i As Long, j As Long
    Dim k As Long
    Dim MapNum As Long
    Dim IDmap() As Long
    Dim NodeMap() As Long
    
    'No data - do nothing
    If NodeIDMax = -1 Then Exit Sub
    
    'if NodeID indexes are too big, we could not use direct mapping
    'max number for direct mapping should be selected with respect to available RAM
    If NodeIDMax > 10000000# Then GoTo lHardWay  'need more than 40M
    
    'SOFT WAY, via direct map from NodeID  (~ O(n) )
    
    ReDim IDmap(NodeIDMax)  'IDmap(NodeID) = index in Nodes array
    
    For i = 0 To NodeIDMax
        IDmap(i) = -1
    Next
    
    For i = 0 To NodesNum - 1
        k = Nodes(i).NodeID
        
        If k < 0 Then GoTo lSkip 'without NodeID - not mergable
        
        If IDmap(k) < 0 Then
            IDmap(k) = i 'first occurence of NodeID
        Else
            Call MergeNodes(IDmap(k), i) 'should join
        End If
lSkip:

        If (i And 8191) = 0 Then
            'display progress
            Form1.Caption = "JoinNodesByID soft " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If

    Next
    GoTo lExit

lHardWay:
    'HARD WAY, via bubble search (~ O(n^2))
    
    ReDim IDmap(NodesNum)  ' IDmap(a) = NodeID
    ReDim NodeMap(NodesNum)  ' NodeMap(a) = index of node in Nodes() array
    MapNum = 0
    
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID >= 0 Then
            For j = 0 To i - 1
                If IDmap(j) = Nodes(i).NodeID Then
                    'found - not first occurence - should join
                    Call MergeNodes(NodeMap(j), i)
                    GoTo lFound
                End If
            Next
            'not found - first occurence of NodeID
            NodeMap(MapNum) = i
            IDmap(MapNum) = Nodes(i).NodeID
            MapNum = MapNum + 1
        End If
lFound:
        
        If (i And 8191) = 0 Then
            'display progress
            Form1.Caption = "JoinNodesByID hard " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next

lExit:
End Sub


'Merge node2 to node1 with relink of all edges
'flag: 1 - ignore node2 coords (0 - move node1 to average coordinates)
Public Sub MergeNodes(node1 As Long, node2 As Long, Optional flag As Long = 0)
    Dim k As Long, i As Long, j As Long
    Dim p As Long
    
    'relink edges from node2 to node1
    p = Nodes(node1).Edges
    k = Nodes(node2).Edges
    For i = 0 To k - 1
        j = Nodes(node2).edge(i)
        
        If Edges(j).node1 = node2 Then
            'edge goes from node2 to X
            Edges(j).node1 = node1
        End If
        If Edges(j).node2 = node2 Then
            'edge goes from X to node2
            Edges(j).node2 = node1
        End If
        'Nodes(node1).edge(p) = j
        Call AddEdgeToNode(node1, j)
        'p = p + 1
    Next
    'Nodes(node1).Edges = p
    Nodes(node2).Edges = 0
    
    'kill all void edges right now
    i = 0
    While i < Nodes(node1).Edges
        j = Nodes(node1).edge(i)
        If Edges(j).node1 = Edges(j).node2 Then Call DelEdge(j)
        i = i + 1
    Wend
    
    If (flag And 1) = 0 Then
        'Calc average coordinates
        'TODO: fix (not safe to 180/-180 edge)
        Nodes(node1).lat = 0.5 * (Nodes(node1).lat + Nodes(node2).lat)
        Nodes(node1).lon = 0.5 * (Nodes(node1).lon + Nodes(node2).lon)
    End If
    
    Call DelNode(node2)
End Sub


'Delete edge and remove all references to it from both nodes
Public Sub DelEdge(ByVal edge1 As Long)
    Dim i As Long
    Dim k As Long
    
    'find this edge among edges of node1
    i = Edges(edge1).node1
    If i = -1 Then Exit Sub 'edge already deleted
    For k = 0 To Nodes(i).Edges - 1
        If Nodes(i).edge(k) = edge1 Then
            'remove edge from edges of node1
            Nodes(i).edge(k) = Nodes(i).edge(Nodes(i).Edges - 1)
            Nodes(i).Edges = Nodes(i).Edges - 1
            GoTo lFound1
        End If
    Next
lFound1:
    
    'find this edge among edges of node2
    i = Edges(edge1).node2
    For k = 0 To Nodes(i).Edges - 1
        If Nodes(i).edge(k) = edge1 Then
            'remove edge from edges of node2
            Nodes(i).edge(k) = Nodes(i).edge(Nodes(i).Edges - 1)
            Nodes(i).Edges = Nodes(i).Edges - 1
            GoTo lFound2
        End If
    Next
lFound2:
    Edges(edge1).node1 = -1 'mark node as deleted
End Sub


'Delete node with all connected edges
Public Sub DelNode(node1 As Long)
    While Nodes(node1).Edges > 0
        Call DelEdge(Nodes(node1).edge(0))
    Wend
    Nodes(node1).Edges = 0
    Nodes(node1).EdgesAlloc = 0
    ReDim Nodes(node1).edge(0)
    Nodes(node1).NodeID = MARK_NODEID_DELETED 'mark node as deleted
End Sub


'Save geometry to simple .mp file (without joining of chains)
Public Sub Save_MP(filename As String)
    Dim i As Long
    Dim k1 As Long, k2 As Long
    Dim typ As Long
    
    Open filename For Output As #2
    Print #2, "; Generated by mp_extsimp"
    Print #2, ""
    'Print #2, MPheader
    
    'custom header
    Print #2, "[IMG ID]"
    Print #2, "CodePage=1251"
    Print #2, "LblCoding=9"
    Print #2, "ID=88888888"
    Print #2, "Name=OSM"
    Print #2, "TypeSet=Navitel"
    Print #2, "Elevation=M"
    Print #2, "Preprocess=F"
    Print #2, "TreSize=3000"
    Print #2, "TreMargin=0.00000"
    Print #2, "RgnLimit=127"
    Print #2, "POIIndex=Y"
    Print #2, "POINumberFirst=N"
    Print #2, "MG=Y"
    Print #2, "Routing=Y"
    Print #2, "Copyright=OpenStreetMap"
    Print #2, "Levels=6"
    Print #2, "Level0=26" 'for high precision of coordinates
    Print #2, "Level1=22"
    Print #2, "Level2=20"
    Print #2, "Level3=18"
    Print #2, "Level4=16"
    Print #2, "Level5=15"
    Print #2, "Zoom0=0"
    Print #2, "Zoom1=1"
    Print #2, "Zoom2=2"
    Print #2, "Zoom3=3"
    Print #2, "Zoom4=4"
    Print #2, "Zoom5=5"
    Print #2, "[END-IMG ID]"
    
    For i = 0 To EdgesNum - 1
        k1 = Edges(i).node1
        k2 = Edges(i).node2
        If k1 < 0 Then GoTo lSkip
        Print #2, "; roadtype=" + CStr(Edges(i).roadtype) + "  edge=" + CStr(i) + " mark=" + CStr(Edges(i).mark)
        Print #2, "[POLYLINE]"
        typ = GetType_by_Highway(Edges(i).roadtype)
        Print #2, "Type=0x"; Hex(typ)
        If Len(Edges(i).label) > 0 Then
            Print #2, "Label=~[0x05]" + Edges(i).label
            Print #2, "StreetDesc=~[0x05]" + Edges(i).label
        End If
        If Edges(i).oneway > 0 Then Print #2, "DirIndicator=1"
        Print #2, "EndLevel=" + CStr(GetTopLevel_by_Highway(Edges(i).roadtype))
        Print #2, "RouteParam=";
        Print #2, CStr(Edges(i).speed); ",";
        Print #2, CStr(GetClass_by_Highway(Edges(i).roadtype)); ",";
        Print #2, CStr(Edges(i).oneway); ",";
        Print #2, "0,0,0,0,0,0,0,0,0"
        Print #2, "Data0=("; Format(Nodes(k1).lat, LATLON_FORMAT); ","; Format(Nodes(k1).lon, LATLON_FORMAT); "),(";
        Print #2, Format(Nodes(k2).lat, LATLON_FORMAT); ","; Format(Nodes(k2).lon, LATLON_FORMAT); ")"
        Print #2, "Nod1=0,"; CStr(k1); ",0"
        Print #2, "Nod2=1,"; CStr(k2); ",0"
        Print #2, "[END]"
        Print #2, ""
    
        If (i And 8191) = 0 Then
            'display progress
            Form1.Caption = "Save_MP " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    
lSkip:
    Next
    
    Close #2
End Sub



'Save geometry to .mp file with joining chains into polylines
Public Sub Save_MP_2(filename As String)
    Dim i As Long
    Dim k1 As Long, k2 As Long
    Dim typ As Long
    
    Open filename For Output As #2
    
    Print #2, "; Generated by mp_extsimp"
    Print #2, ""
    Print #2, MPheader
    
    For i = 0 To EdgesNum - 1
        If Edges(i).node1 = -1 Then
            'deleted edge
            Edges(i).mark = 1 'mark to ignore
        Else
            Edges(i).mark = 0 'mark to save
        End If
    Next
    
    For i = 0 To EdgesNum - 1
        If Edges(i).mark = 0 Then
            'all marked to save - find chain and save
            Call SaveChain(i)
        End If
        
        If (i And 8191) = 0 Then
            'display progress
            Form1.Caption = "Save_MP_2 " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    Next
    
    Print #2, "; Completed" 'file finalization flag
    
    Close #2
End Sub



'Find and optimize all chains by Douglas-Peucker with Epsilon (in metres)
Public Sub DouglasPeucker_total(Epsilon As Double)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To NodesNum - 1
        Nodes(i).mark = 0 'mark all nodes as not passed
    Next
    
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID = MARK_NODEID_DELETED Or Nodes(i).Edges <> 2 Or Nodes(i).mark = 1 Then GoTo lSkip
            'node: not deleted, not yet passed and with 2 edges -> should be checked for chain
            Call DouglasPeucker_chain(i, Epsilon)
lSkip:
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "DouglasPeucker_total (" + CStr(Epsilon) + ") " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next
End Sub


'find and optimize all chains by Douglas-Peucker with Epsilon (in metres) and limiting max edge (in metres)
Public Sub DouglasPeucker_total_split(Epsilon As Double, MaxEdge As Double)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To NodesNum - 1
        Nodes(i).mark = 0 'mark all nodes as not passed
    Next
    
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID = MARK_NODEID_DELETED Or Nodes(i).Edges <> 2 Or Nodes(i).mark = 1 Then GoTo lSkip
            'node: not deleted, not yet passed and with 2 edges -> should be checked for chain
            Call DouglasPeucker_chain_split(i, Epsilon, MaxEdge)
lSkip:
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "DouglasPeucker_total_split (" + CStr(Epsilon) + ", " + CStr(MaxEdge) + ") " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next
End Sub



'Find one chain (starting from node1) and optimize it by Douglas-Peucker with Epsilon (in metres)
Public Sub DouglasPeucker_chain(node1 As Long, Epsilon As Double)
    Dim i As Long
    Dim j As Long
    Dim k As Long, m As Long
    Dim refedge As edge
    Dim ChainEnd As Long
    Dim NextChainEdge As Long
    Dim node0 As Long
    
    NextChainEdge = -1
    ChainEnd = 0
    
    'Algorithm go from specified node into one direction by chain of nodes
    '(nodes connected one by one, without junctions) until end (or junction) is reached
    'After that algorithm will go from final edge into opposite direction and will compare edges
    'and add nodes into Chain array
    'On findind different edge (or reaching other end of chain) algorithm will pass found (sub)chain
    'into OptimizeByDouglasPeucker_One recursive function for optimization
    'Then rest of chain (if it exits) will be processed in similar way
    
    '1) go by chain to the one end - to node with !=2 edges
    
    i = node1 'start node
    j = node1
lGoNext:
    k = GoByChain(i, j) 'go by chain
    If k <> node1 And Nodes(k).Edges = 2 Then j = i: i = k: GoTo lGoNext 'if still 2 edges - proceed
    
    
    '   *-----*-----*-----*---...
    '   k     i     j
    
    j = k 'OK, we found end of chain
    
    '   *---------*-----*-----*---...
    '  k=j        i
    
    '2) go revert - from found end to another one and saving all nodes into Chain() array
    
    ChainNum = 0
    node0 = k 'start node
    Call AddChain(k)
    Call AddChain(i)
    
    'keep info about first edge in chain
    refedge = Edges(GoByChain_lastedge)
    If refedge.node1 <> Chain(0) And refedge.oneway = 1 Then refedge.oneway = 2 'reversed oneway

lGoNext2:

    k = GoByChain(i, j)
    
    '   *-------------*-----*-----*---...
    '  j              i     k
    
    'check oneway
    m = Edges(GoByChain_lastedge).oneway
    If m > 0 And Edges(GoByChain_lastedge).node1 <> i Then m = 2
    
    'if oneway flag is differnt or road type is changed - break chain
    If m <> refedge.oneway Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    If Edges(GoByChain_lastedge).roadtype <> refedge.roadtype Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    
    Call AddChain(k)
    
    If k <> Chain(0) And k <> node0 And Nodes(k).Edges = 2 Then
        'still 2 edges - still chain, found first node from chain or node0 - chain loop, exit
        Nodes(k).mark = 1
        j = i
        i = k
        GoTo lGoNext2
    End If
    
    ChainEnd = 1
    
lBreak:

    '3) optimize found chain by D-P
    Call OptimizeByDouglasPeucker_One(0, ChainNum - 1, Epsilon, refedge)
    
    If ChainEnd = 0 Then
        'continue with this chain, as it is not ended
        
        '   *================*--------------------*-----------*-----*---...
        '                        NextChainEdge
        
        'new reference info
        refedge = Edges(NextChainEdge)
        If refedge.node1 = Chain(ChainNum - 1) Then
            i = refedge.node2
            j = refedge.node1
        Else
            If refedge.oneway = 1 Then refedge.oneway = 2
            i = refedge.node1
            j = refedge.node2
        End If
        
        '   *================*--------------------*-----------*-----*---...
        '                    j                    i
        
        If Nodes(i).Edges <> 2 Or j = node0 Or i = node0 Then Exit Sub 'chain from one edge or loop found - nothing to optimize by D-P
        
        'add both nodes of last edge
        ChainNum = 0
        Call AddChain(j)
        Call AddChain(i)
        
        NextChainEdge = -1
        GoTo lGoNext2 'continue with chain
    End If

End Sub


'find one chain (starting from node1) and optimize it by Douglas-Peucker with Epsilon (in metres) and limiting edge len by MaxEdge
Public Sub DouglasPeucker_chain_split(node1 As Long, Epsilon As Double, MaxEdge As Double)
    Dim i As Long
    Dim j As Long
    Dim k As Long, m As Long
    Dim refedge As edge
    Dim ChainEnd As Long
    Dim NextChainEdge As Long
    Dim node0 As Long
    
    NextChainEdge = -1
    ChainEnd = 0
    
    'Algorithm works as DouglasPeucker_chain above
    'difference is only inside OptimizeByDouglasPeucker_One_split
    
    '1) go by chain to the one end - to node with !=2 edges
    
    i = node1 'start node
    j = node1
lGoNext:
    k = GoByChain(i, j) 'go by chain
    If k <> node1 And Nodes(k).Edges = 2 Then j = i: i = k: GoTo lGoNext 'if still 2 edges - proceed
    
    '   *-----*-----*-----*---...
    '   k     i     j
    
    j = k 'OK, we found end of chain
    
    '   *---------*-----*-----*---...
    '  k=j        i
    
    '2) go revert - from found end to another one and saving all nodes into Chain() array
    
    ChainNum = 0
    node0 = k 'start node
    Call AddChain(k)
    Call AddChain(i)
    
    'keep info about first edge in chain
    refedge = Edges(GoByChain_lastedge)
    If refedge.node1 <> Chain(0) And refedge.oneway = 1 Then refedge.oneway = 2 'reversed oneway

lGoNext2:

    k = GoByChain(i, j)
    
    '   *-------------*-----*-----*---...
    '  j              i     k
    
    'check oneway
    m = Edges(GoByChain_lastedge).oneway
    If m > 0 And Edges(GoByChain_lastedge).node1 <> i Then m = 2
    
    'if oneway flag is differnt or road type is changed - break chain
    If m <> refedge.oneway Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    If Edges(GoByChain_lastedge).roadtype <> refedge.roadtype Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    
    Call AddChain(k)
    
    'If k <> Chain(0) And Nodes(k).Edges = 2 And Nodes(k).mark = 0 Then
    If k <> Chain(0) And k <> node0 And Nodes(k).Edges = 2 Then
        'still 2 edges - still chain, found first node from chain or node0 - chain loop, exit
        Nodes(k).mark = 1
        j = i
        i = k
        GoTo lGoNext2
    End If
    
    ChainEnd = 1
    
lBreak:

    '3) optimize found chain by D-P
    Call OptimizeByDouglasPeucker_One_split(0, ChainNum - 1, Epsilon, refedge, MaxEdge)
    
    If ChainEnd = 0 Then
        'continue with this chain, as it is not ended
        
        '   *================*--------------------*-----------*-----*---...
        '                        NextChainEdge
        
        'new reference info
        refedge = Edges(NextChainEdge)
        If refedge.node1 = Chain(ChainNum - 1) Then
            i = refedge.node2
            j = refedge.node1
        Else
            If refedge.oneway = 1 Then refedge.oneway = 2
            i = refedge.node1
            j = refedge.node2
        End If
        
        '   *================*--------------------*-----------*-----*---...
        '                    j                    i
        
        If Nodes(i).Edges <> 2 Or j = node0 Or i = node0 Then Exit Sub 'chain from one edge or loop found - nothing to optimize by D-P
        
        'add both nodes of last edge
        ChainNum = 0
        Call AddChain(j)
        Call AddChain(i)

        NextChainEdge = -1
        GoTo lGoNext2 'continue with chain
    End If
    

End Sub


'Save chain of edges into mp file (already opened as #2)
Public Sub SaveChain(edge1 As Long)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long, m As Long
    Dim refedge As edge
    Dim ChainEnd As Long
    Dim NextChainEdge As Long
    Dim startnode As Long
    
    Dim k1 As Long, k2 As Long
    Dim typ As Long
    
    NextChainEdge = -1
    ChainEnd = 0
    
    'Algorithm go from specified edge into one direction by chain of nodes
    '(nodes connected one by one, without junctions) until end (or junction) is reached
    'After that algorithm will go from final edge into opposite direction and will compare edges
    'and add nodes into Chain array
    'On findind different edge (or reaching other end of chain) algorithm will save found (sub)chain into mp file
    'Then rest of chain (if it exits) will be processed in similar way
    
    '1) go by chain to the one end - to node with !=2 edges
    
    i = Edges(edge1).node1 'start node
    j = Edges(edge1).node2
    
    startnode = j
    
    If Nodes(i).Edges <> 2 Then
        'i is end of chain
        ChainNum = 0
        Call AddChain(i)
        Call AddChain(j)
        startnode = i 'for detecting loops
        refedge = Edges(edge1)
        
        Edges(edge1).mark = 1 'saved
        
        If Nodes(j).Edges <> 2 Then
            'that's all
            ChainEnd = 1
            GoTo lBreak
        Else
            j = Chain(0)
            i = Chain(1)
            GoTo lGoNext2
        End If
    End If
    
lGoNext:
    k = GoByChain(i, j) 'go by chain
    If Nodes(k).Edges = 2 And k <> startnode Then j = i: i = k: GoTo lGoNext 'if still 2 edges and we have not found loop - proceed
    
    '   *-----*-----*-----*---...
    '   k     i     j
    
    j = k 'OK, we found end of chain
    
    '   *---------*-----*-----*---...
    '  k=j        i
    
    '2) go revert - from found end to another one and saving all nodes into Chain() array
    
    ChainNum = 0
    Call AddChain(k)
    Call AddChain(i)
    startnode = k
    
    'keep info about first edge in chain
    refedge = Edges(GoByChain_lastedge)
    Edges(GoByChain_lastedge).mark = 1
    If refedge.node1 <> Chain(0) And refedge.oneway = 1 Then refedge.oneway = 2 'reversed oneway

lGoNext2:

    k = GoByChain(i, j)
    
    '   *-------------*-----*-----*---...
    '  j              i     k
    
    'check oneway
    m = Edges(GoByChain_lastedge).oneway
    If m > 0 And Edges(GoByChain_lastedge).node1 <> i Then m = 2
    
    'if oneway flag is differnt or road type, speed or label is changed - break chain
    If m <> refedge.oneway Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    If Edges(GoByChain_lastedge).roadtype <> refedge.roadtype Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    If Edges(GoByChain_lastedge).speed <> refedge.speed Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    If Edges(GoByChain_lastedge).label <> refedge.label Then NextChainEdge = GoByChain_lastedge: GoTo lBreak
    
    Edges(GoByChain_lastedge).mark = 1 'saved
    
    Call AddChain(k)
    
    If Nodes(k).Edges = 2 And k <> startnode Then
        'still 2 edges - still chain
        j = i
        i = k
        GoTo lGoNext2
    End If
    
    ChainEnd = 1
    
lBreak:

    '3) save chain to file

    'Print #2, "; roadtype=" + CStr(refedge.roadtype) 'debug info about road type
    Print #2, "[POLYLINE]"
    typ = GetType_by_Highway(refedge.roadtype)  'object type - from road type
    Print #2, "Type=0x"; Hex(typ)
    If Len(refedge.label) > 0 Then
        'labels - into special codes fro labelization
        Print #2, "Label=~[0x05]" + refedge.label
        Print #2, "StreetDesc=~[0x05]" + refedge.label
    End If
    If refedge.oneway > 0 Then Print #2, "DirIndicator=1" 'oneway indicator
    Print #2, "EndLevel=" + CStr(GetTopLevel_by_Highway(refedge.roadtype)) 'top level of visibility - from road type
    Print #2, "RouteParam=";
    Print #2, CStr(refedge.speed); ","; 'speed class
    Print #2, CStr(GetClass_by_Highway(refedge.roadtype)); ","; 'road class - from road type
    If refedge.oneway > 0 Then
        Print #2, "1,"; 'one_way
    Else
        Print #2, "0,";
    End If
    Print #2, "0,0,0,0,0,0,0,0,0" 'other params are not handled
    Print #2, "Data0=";
    If refedge.oneway = 2 Then
        'reverted oneway, save in backward sequence
        For i = ChainNum - 1 To 0 Step -1
            If i <> ChainNum - 1 Then Print #2, ",";
            Print #2, "("; Format(Nodes(Chain(i)).lat, LATLON_FORMAT); ","; Format(Nodes(Chain(i)).lon, LATLON_FORMAT); ")";
        Next
        Print #2,
        Print #2, "Nod1=0,"; CStr(Chain(ChainNum - 1)); ",0"
        Print #2, "Nod2=" + CStr(ChainNum - 1) + ","; CStr(Chain(0)); ",0"
    Else
        'forward oneway or twoway, save in direct sequence
        For i = 0 To ChainNum - 1
            If i <> 0 Then Print #2, ",";
            Print #2, "("; Format(Nodes(Chain(i)).lat, LATLON_FORMAT); ","; Format(Nodes(Chain(i)).lon, LATLON_FORMAT); ")";
        Next
        Print #2,
        Print #2, "Nod1=0,"; CStr(Chain(0)); ",0"
        Print #2, "Nod2=" + CStr(ChainNum - 1) + ","; CStr(Chain(ChainNum - 1)); ",0"
    End If
    Print #2, "[END]"
    Print #2, ""

    If ChainEnd = 0 Then
        'continue with this chain, as it is not ended
        
        '   *================*--------------------*-----------*-----*---...
        '                        NextChainEdge
        
        'new reference info
        refedge = Edges(NextChainEdge)
        If refedge.node1 = Chain(ChainNum - 1) Then
            i = refedge.node2
            j = refedge.node1
        Else
            If refedge.oneway = 1 Then refedge.oneway = 2
            i = refedge.node1
            j = refedge.node2
        End If
        
        '   *================*--------------------*-----------*-----*---...
        '                    j                    i
        
        Edges(NextChainEdge).mark = 1
        
        If j = startnode Or i = startnode Then Exit Sub 'loop detected - exit
        'If j = startnode Then Print #2, "; startnode"
        
        'add both nodes of last edge
        ChainNum = 0
        Call AddChain(j)
        Call AddChain(i)
        
        If Nodes(i).Edges <> 2 Then
            'chain from one edge
            ChainEnd = 1
            GoTo lBreak
        End If
        
        NextChainEdge = -1
        GoTo lGoNext2 'continue with chain
    End If
    

End Sub


'Go by chain from node1 in some direction, but not to Node0
'(assumed, that node1 have two edges, not 1, not 3 or more, otherwise - UB)
'Usage: GoByChain(x,x) goes by first edge, z=GoByChain(x,y)->u=GoByChain(z,x)->... allows to travel by chain node by node
Public Function GoByChain(node1 As Long, node0 As Long) As Long
    Dim i As Long, k As Long
    
    'check first edge
    i = Nodes(node1).edge(0)
    k = Edges(i).node1
    If k = node1 Then k = Edges(i).node2
    GoByChain_lastedge = i
    If k = node0 Then
        'node0 -> check second edge
        i = Nodes(node1).edge(1)
        k = Edges(i).node1
        If k = node1 Then k = Edges(i).node2
        GoByChain_lastedge = i
    End If
    GoByChain = k
End Function

'Get edge, conecting node1 and node2, return -1 if no connection
'TODO(opt): swap node1 and node2 if node2 have smaller edges
Public Function GetEdgeBetween(node1 As Long, node2 As Long) As Long
    Dim i As Long, j As Long
    For i = 0 To Nodes(node1).Edges - 1
        j = Nodes(node1).edge(i)
        If Edges(j).node1 = node2 Or Edges(j).node2 = node2 Then
            'found
            GetEdgeBetween = j
            Exit Function
        End If
    Next
    GetEdgeBetween = -1
End Function

'Parse OSM highway class to our own constants
Public Function GetHighwayType(Text As String) As Long
    Select Case LCase(Trim(Text))
        Case "primary"
            GetHighwayType = HIGHWAY_PRIMARY
        Case "primary_link"
            GetHighwayType = HIGHWAY_PRIMARY_LINK
        Case "secondary"
            GetHighwayType = HIGHWAY_SECONDARY
        Case "secondary_link"
            GetHighwayType = HIGHWAY_SECONDARY_LINK
        Case "tertiary"
            GetHighwayType = HIGHWAY_TERTIARY
        Case "tertiary_link"
            GetHighwayType = HIGHWAY_TERTIARY_LINK
        Case "motorway"
            GetHighwayType = HIGHWAY_MOTORWAY
        Case "motorway_link"
            GetHighwayType = HIGHWAY_MOTORWAY_LINK
        Case "trunk"
            GetHighwayType = HIGHWAY_TRUNK
        Case "trunk_link"
            GetHighwayType = HIGHWAY_TRUNK_LINK
        Case "living_street"
            GetHighwayType = HIGHWAY_LIVING_STREET
        Case "residential"
            GetHighwayType = HIGHWAY_RESIDENTIAL
        Case "unclassified"
            GetHighwayType = HIGHWAY_UNCLASSIFIED
        Case "service"
            GetHighwayType = HIGHWAY_SERVICE
        Case "track"
            GetHighwayType = HIGHWAY_TRACK
        Case "road"
            GetHighwayType = HIGHWAY_UNKNOWN
        Case Else
            GetHighwayType = HIGHWAY_OTHER
    End Select
End Function


'Convert constants to polyline type
Public Function GetType_by_Highway(ByVal HighwayType As Long) As Long
    Select Case HighwayType
        Case HIGHWAY_MOTORWAY
            GetType_by_Highway = 1
        Case HIGHWAY_MOTORWAY_LINK
            GetType_by_Highway = 9
        Case HIGHWAY_TRUNK
            GetType_by_Highway = Control_TrunkType
        Case HIGHWAY_TRUNK_LINK
            GetType_by_Highway = Control_TrunkLinkType
        Case HIGHWAY_PRIMARY
            GetType_by_Highway = Control_PrimaryType
        Case HIGHWAY_PRIMARY_LINK
            GetType_by_Highway = 8
        Case HIGHWAY_SECONDARY
            GetType_by_Highway = 3
        Case HIGHWAY_SECONDARY_LINK
            GetType_by_Highway = 8
        Case HIGHWAY_TERTIARY
            GetType_by_Highway = 3
        Case HIGHWAY_TERTIARY_LINK
            GetType_by_Highway = 8
        Case HIGHWAY_LIVING_STREET
            GetType_by_Highway = 6
        Case HIGHWAY_RESIDENTIAL
            GetType_by_Highway = 6
        Case HIGHWAY_UNCLASSIFIED
            GetType_by_Highway = 3
        Case HIGHWAY_SERVICE
            GetType_by_Highway = 7
        Case HIGHWAY_TRACK
            GetType_by_Highway = 10
        Case HIGHWAY_UNKNOWN
            GetType_by_Highway = 3
        Case HIGHWAY_OTHER
            GetType_by_Highway = 3
        Case Else
            GetType_by_Highway = 3
    End Select
End Function


'Convert constants to road class
Public Function GetClass_by_Highway(ByVal HighwayType As Long) As Long
    Select Case HighwayType
        Case HIGHWAY_MOTORWAY
            GetClass_by_Highway = 4
        Case HIGHWAY_MOTORWAY_LINK
            GetClass_by_Highway = 4
        Case HIGHWAY_TRUNK
            GetClass_by_Highway = 4
        Case HIGHWAY_TRUNK_LINK
            GetClass_by_Highway = 4
        Case HIGHWAY_PRIMARY
            GetClass_by_Highway = 3
        Case HIGHWAY_PRIMARY_LINK
            GetClass_by_Highway = 3
        Case HIGHWAY_SECONDARY
            GetClass_by_Highway = 2
        Case HIGHWAY_SECONDARY_LINK
            GetClass_by_Highway = 2
        Case HIGHWAY_TERTIARY
            GetClass_by_Highway = 1
        Case HIGHWAY_TERTIARY_LINK
            GetClass_by_Highway = 1
        Case HIGHWAY_LIVING_STREET
            GetClass_by_Highway = 0
        Case HIGHWAY_RESIDENTIAL
            GetClass_by_Highway = 0
        Case HIGHWAY_UNCLASSIFIED
            GetClass_by_Highway = 1
        Case HIGHWAY_SERVICE
            GetClass_by_Highway = 0
        Case HIGHWAY_TRACK
            GetClass_by_Highway = 0
        Case HIGHWAY_UNKNOWN
            GetClass_by_Highway = 0
        Case HIGHWAY_OTHER
            GetClass_by_Highway = 0
        Case Else
            GetClass_by_Highway = 0
    End Select
End Function


'Convert constants to top level for visibility
Public Function GetTopLevel_by_Highway(ByVal HighwayType As Long) As Long
    Select Case HighwayType
        Case HIGHWAY_MOTORWAY
            GetTopLevel_by_Highway = 6
        Case HIGHWAY_MOTORWAY_LINK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_TRUNK
            GetTopLevel_by_Highway = 6
        Case HIGHWAY_TRUNK_LINK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_PRIMARY
            GetTopLevel_by_Highway = 5
        Case HIGHWAY_PRIMARY_LINK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_SECONDARY
            GetTopLevel_by_Highway = 4
        Case HIGHWAY_SECONDARY_LINK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_TERTIARY
            GetTopLevel_by_Highway = 3
        Case HIGHWAY_TERTIARY_LINK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_LIVING_STREET
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_RESIDENTIAL
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_UNCLASSIFIED
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_SERVICE
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_TRACK
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_UNKNOWN
            GetTopLevel_by_Highway = 2
        Case HIGHWAY_OTHER
            GetTopLevel_by_Highway = 2
        Case Else
            GetTopLevel_by_Highway = 2
    End Select
End Function



'Move node 3 to closest point on line node1-node2
'TODO: fix (not safe to 180/-180 edge)
Public Sub ProjectNode(node1 As Long, node2 As Long, ByRef node3 As node)
    Dim k As Double
    Dim Xab As Double
    Dim Yab As Double
    Xab = Nodes(node2).lat - Nodes(node1).lat
    Yab = Nodes(node2).lon - Nodes(node1).lon
    k = (Xab * (node3.lat - Nodes(node1).lat) + Yab * (node3.lon - Nodes(node1).lon)) / (Xab * Xab + Yab * Yab)
    node3.lat = Nodes(node1).lat + k * Xab
    node3.lon = Nodes(node1).lon + k * Yab
End Sub


'Calc cosine of angle betweeen two edges
'(calc via vectors on reference ellipsoid, 180/-180 safe)
Public Function CosAngleBetweenEdges(edge1 As Long, edge2 As Long) As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim x3 As Double, y3 As Double, z3 As Double
    Dim x4 As Double, y4 As Double, z4 As Double
    Dim node1 As Long, node2 As Long
    Dim len1 As Double, len2 As Double
    
    'XYZ
    node1 = Edges(edge1).node1
    node2 = Edges(edge1).node2
    Call LatLonToXYZ(Nodes(node1).lat, Nodes(node1).lon, x1, y1, z1)
    Call LatLonToXYZ(Nodes(node2).lat, Nodes(node2).lon, x2, y2, z2)
    node1 = Edges(edge2).node1
    node2 = Edges(edge2).node2
    Call LatLonToXYZ(Nodes(node1).lat, Nodes(node1).lon, x3, y3, z3)
    Call LatLonToXYZ(Nodes(node2).lat, Nodes(node2).lon, x4, y4, z4)
    
    'vectors
    x2 = x2 - x1
    y2 = y2 - y1
    z2 = z2 - z1
    x4 = x4 - x3
    y4 = y4 - y3
    z4 = z4 - z3
    
    'vector lengths
    len1 = Sqr(x2 * x2 + y2 * y2 + z2 * z2)
    len2 = Sqr(x4 * x4 + y4 * y4 + z4 * z4)
    
    If len1 = 0 Or len2 = 0 Then
        'one of vectors is void
        CosAngleBetweenEdges = 0
    Else
        'Cosine of angle is scalar multiply divided by lengths
        CosAngleBetweenEdges = (x2 * x4 + y2 * y4 + z2 * z4) / (len1 * len2)
    End If

End Function


'Convert (lat,lon) to (x,y,z) on reference ellipsoid
Public Function LatLonToXYZ(lat As Double, lon As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    Dim r As Double
    r = DATUM_R_EQUAT * Cos(lat * DEGTORAD)
    z = DATUM_R_POLAR * Sin(lat * DEGTORAD)
    x = r * Sin(lon * DEGTORAD)
    y = r * Cos(lon * DEGTORAD)
End Function


'Calc distance square from node1 to node2 in metres
'metric distance of ellipsoid chord (not arc)
Public Function DistanceSquare(node1 As Long, node2 As Long) As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Call LatLonToXYZ(Nodes(node1).lat, Nodes(node1).lon, x1, y1, z1)
    Call LatLonToXYZ(Nodes(node2).lat, Nodes(node2).lon, x2, y2, z2)
    DistanceSquare = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2) + (z1 - z2) * (z1 - z2)
End Function


'Calc distance from node1 to node2 in metres
Public Function Distance(node1 As Long, node2 As Long) As Double
    Distance = Sqr(DistanceSquare(node1, node2))
End Function



'Calc distance from node3 to interval (not just line) from node1 to node2 in metres
'Calc by Heron's formula from sides of triangle   (180/-180 safe)
Public Function DistanceToSegment(node1 As Long, node2 As Long, node3 As Long) As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim s2 As Double
    
    a = DistanceSquare(node1, node2) 'Calc squares of triangle sides
    b = DistanceSquare(node1, node3)
    c = DistanceSquare(node2, node3)
    If a = 0 Then
        DistanceToSegment = Sqr(b)
        DistanceToSegment_last_case = 0 'node1=node2
        Exit Function
    ElseIf b > (a + c) Then
        DistanceToSegment = Sqr(c)
        DistanceToSegment_last_case = 1 'node1 is closest point to node3
        Exit Function
    ElseIf c > (a + b) Then
        DistanceToSegment = Sqr(b)
        DistanceToSegment_last_case = 2 'node2 is closest point to node3
        Exit Function
    Else
        a = Sqr(a) 'Calc sides lengths from squares
        b = Sqr(b)
        c = Sqr(c)
        s2 = 0.5 * Sqr((a + b + c) * (a + b - c) * (a + c - b) * (b + c - a))
        DistanceToSegment = s2 / a
        DistanceToSegment_last_case = 3 'closest point is inside interval
    End If
End Function


'Calc distance between not crossing edges (edge1 and edge2)
'(180/-180 safe)
Public Function DistanceBetweenSegments(edge1 As Long, edge2 As Long) As Double
    Dim d1 As Double, d2 As Double
    'Just minimum of 4 distances (each ends to each other edge)
    d1 = DistanceToSegment(Edges(edge1).node1, Edges(edge1).node2, Edges(edge2).node1)
    d2 = DistanceToSegment(Edges(edge1).node1, Edges(edge1).node2, Edges(edge2).node2)
    If d2 < d1 Then d1 = d2
    d2 = DistanceToSegment(Edges(edge2).node1, Edges(edge2).node2, Edges(edge1).node1)
    If d2 < d1 Then d1 = d2
    d2 = DistanceToSegment(Edges(edge2).node1, Edges(edge2).node2, Edges(edge1).node2)
    If d2 < d1 Then d1 = d2
    DistanceBetweenSegments = d1
End Function


'Recursive check to optimize chain/subchain by Douglas-Peucker with Epsilon (in metres)
'subchain is defined by IndexStart,IndexLast
'refedge - road parameters of chain (for create new edge in case of optimization)
'(180/-180 safe)
Private Sub OptimizeByDouglasPeucker_One(IndexStart As Long, IndexLast As Long, Epsilon As Double, refedge As edge)
    Dim i As Long
    Dim FarestIndex As Long
    Dim FarestDist As Double
    Dim Dist As Double
    Dim k As Double
    Dim ScalarMult As Double
    Dim newspeed As Integer
    Dim newlabel As String
    
    If ((IndexStart + 1) >= IndexLast) Then Exit Sub 'one edge (or less) -> nothing to do

    k = Distance(Chain(IndexStart), Chain(IndexLast)) 'distance between subchain edge

    'find node, farest from line first-last node (farer than Epsilon)
    FarestDist = Epsilon 'start max len - Epsilon
    FarestIndex = -1 'nothing yet found
    For i = IndexStart + 1 To IndexLast - 1
        If k = 0 Then
            'circled subchain
            Dist = Distance(Chain(i), Chain(IndexStart))
        Else
            Dist = DistanceToSegment(Chain(IndexStart), Chain(IndexLast), Chain(i))
        End If
        If Dist > FarestDist Then
            FarestDist = Dist: FarestIndex = i
        End If
    Next

    If FarestIndex = -1 Then
        'farest node not found -> all distances less than Epsilon -> remove all internal nodes
        
        'calc speed and label from all subchain edges
        Call EstimateChain(IndexStart, IndexLast)
        newspeed = EstimateChain_speed
        newlabel = EstimateChain_label
        
        For i = IndexStart + 1 To IndexLast - 1
            Call DelNode(Chain(i)) 'kill all nodes with edges
        Next
        
        'join first and last nodes by new edge
        If refedge.oneway = 2 Then
            'reversed oneway
            i = JoinByEdge(Chain(IndexLast), Chain(IndexStart))
            Edges(i).oneway = 1
        Else
            i = JoinByEdge(Chain(IndexStart), Chain(IndexLast))
            Edges(i).oneway = refedge.oneway
        End If
        Edges(i).roadtype = refedge.roadtype
        Edges(i).speed = newspeed
        Edges(i).label = newlabel
        
        Exit Sub
    End If

    'farest point found - keep it
    'call Douglas-Peucker for two new subchains
    Call OptimizeByDouglasPeucker_One(IndexStart, FarestIndex, Epsilon, refedge)
    Call OptimizeByDouglasPeucker_One(FarestIndex, IndexLast, Epsilon, refedge)
End Sub


'Recursive check to optimize chain/subchain by Douglas-Peucker with Epsilon (in metres) and limiting edge len by MaxEdge
'subchain is defined by IndexStart,IndexLast
'refedge - road parameters of chain (for create new edge in case of optimization)
'(180/-180 safe)
Private Sub OptimizeByDouglasPeucker_One_split(IndexStart As Long, IndexLast As Long, Epsilon As Double, refedge As edge, MaxEdge As Double)
    Dim i As Long
    Dim FarestIndex As Long
    Dim FarestDist As Double
    Dim Dist As Double
    Dim k As Double
    Dim ScalarMult As Double
    Dim newspeed As Integer
    Dim newlabel As String
    
    If ((IndexStart + 1) >= IndexLast) Then Exit Sub 'one edge (or less) -> nothing to do

    k = Distance(Chain(IndexStart), Chain(IndexLast)) 'distance between subchain edge

    'find node, farest from line first-last node (farer than Epsilon)
    FarestDist = Epsilon 'start max len - Epsilon
    FarestIndex = -1 'nothing yet found
    For i = IndexStart + 1 To IndexLast - 1
        If k = 0 Then
            'circled subchain
            Dist = Distance(Chain(i), Chain(IndexStart))
        Else
            Dist = DistanceToSegment(Chain(IndexStart), Chain(IndexLast), Chain(i))
        End If
        If Dist > FarestDist Then
            FarestDist = Dist: FarestIndex = i
        End If
        
        If Distance(Chain(i), Chain(IndexStart)) > MaxEdge Then
            'distance from start to this node is more than limit -> we should keep this node
            FarestIndex = i
            GoTo lKeepFar
        End If
    Next

    If FarestIndex = -1 Then
        'farest node not found -> all distances less than Epsilon -> remove all internal nodes
        
        'calc speed and label from all subchain edges
        Call EstimateChain(IndexStart, IndexLast)
        newspeed = EstimateChain_speed
        newlabel = EstimateChain_label
        
        For i = IndexStart + 1 To IndexLast - 1
            Call DelNode(Chain(i)) 'kill with edges
        Next
        
        'join first and last nodes by new edge
        If refedge.oneway = 2 Then
            'reversed oneway
            i = JoinByEdge(Chain(IndexLast), Chain(IndexStart))
            Edges(i).oneway = 1
        Else
            i = JoinByEdge(Chain(IndexStart), Chain(IndexLast))
            Edges(i).oneway = refedge.oneway
        End If
        Edges(i).roadtype = refedge.roadtype
        Edges(i).speed = newspeed
        Edges(i).label = newlabel
        
        Exit Sub
    End If

lKeepFar:
    'farest point found - keep it
    'call Douglas-Peucker for two new subchains
    Call OptimizeByDouglasPeucker_One_split(IndexStart, FarestIndex, Epsilon, refedge, MaxEdge) 'Douglas-Peucker for two new subchains
    Call OptimizeByDouglasPeucker_One_split(FarestIndex, IndexLast, Epsilon, refedge, MaxEdge)

End Sub


'Load .mp file
'Remove _link flags from polylines longer than MaxLinkLen
'(loader is basic and rather stupid, uses relocation on file to read section info without internal buffering)
Public Sub Load_MP(filename As String, MaxLinkLen As Double)
    Dim LogOptimization As Long
    Dim sLine As String
    Dim fLat As Double
    Dim fLon As Double
    Dim sWay As String
    Dim SectionType As Long '0 - none, 1 - header, 2 - polyline, 3 - polygon
    Dim iPhase As Long  'Phase of reading polyline: 0 - general part, 1 - scan for routeparam, 2 - scan for geometry, 3 - scan for routing (nodeid e.t.c)
    Dim iStartLine As Long
    Dim iPrevLine As Long
    Dim FileLen As Long
    Dim sPrefix As String
    Dim DataLineNum As Long
    Dim k As Long, k2 As Long, k3 As Long, p As Long
    Dim i As Long, j As Long
    Dim ThisLineNodes As Long
    Dim NodeID As Long
    Dim WayClass As Long
    Dim WaySpeed As Long
    Dim WayOneway As Long
    Dim routep() As String
    Dim LastCommentHighway As Long
    Dim label As String
    Dim LinkLen As Double
    Dim LastWayID As String
    Dim NumDelinked As Long
    
    NodeIDMax = -1 'no nodeid yet
    
    Open filename For Input As #1
    FileLen = LOF(1)
    
    SectionType = 0
    WayClass = -1
    WaySpeed = -1
    iPhase = 0
    label = ""
    iStartLine = 0
    iPrevLine = 0
    MPheader = ""
    LastCommentHighway = HIGHWAY_UNSPECIFIED
    LastWayID = ""
    NumDelinked = 0
    
lNextLine:
    iPrevLine = Seek(1) 'get current position in file
    Line Input #1, sLine
    
    'check for section start
    If InStr(1, sLine, "[IMG ID]") > 0 Then
        'header section
        SectionType = 1
        iPhase = 0
    End If
    If InStr(1, sLine, "[POLYGON]") > 0 Then
        'polygon
        SectionType = 3
        GoTo lStartPoly
    End If
    If InStr(1, sLine, "[POLYLINE]") > 0 Then
        'polyline
        SectionType = 2
lStartPoly:
        If (iPrevLine And 1023) = 0 Then
            'display progress
            Form1.Caption = "Load_MP (" + CStr(MaxLinkLen) + ") " + CStr(iPrevLine) + " / " + CStr(FileLen): Form1.Refresh
        End If
        DataLineNum = 0
        If iPhase = 0 Then
            'first pass of section? start scanning
            WayClass = -1
            WaySpeed = -1
            iPhase = 1
            iStartLine = iPrevLine 'remember current pos (where to go after ending pass)
        End If
    End If
    
    If InStr(1, sLine, "[END") > 0 Then
        'section ended
        
        If SectionType = 1 Then
            MPheader = MPheader + sLine + vbNewLine 'add ending of section into saved header
        End If
        
        
        If iPhase = 1 And WaySpeed = -1 Then
            If Control_LoadNoRoute = 1 Then
                If WaySpeed = -1 Then WaySpeed = Control_ForceWaySpeed
                If WayClass = -1 Then WayClass = HIGHWAY_SECONDARY
            Else
                iPhase = 0 'no routing params found in 1st pass - skip way completely
            End If
        End If
            

        If iPhase > 0 And iPhase < 3 Then
            'not last pass of section -> goto start of it
            Seek 1, iStartLine 'relocate in file
            iPhase = iPhase + 1
            GoTo lNextLine
        End If
        
        LastCommentHighway = HIGHWAY_UNSPECIFIED 'if no osm2mp info yet found
        label = ""
        iPhase = 0
        SectionType = 0
    End If
    
    Select Case iPhase
        Case 0
            If Left(Trim(sLine), 9) = "; highway" Then
                'comment, produced by osm2mp
                LastCommentHighway = GetHighwayType(Trim(Mid(sLine, 12, Len(sLine))))
            End If
            If Left(Trim(sLine), 7) = "; WayID" Then
                'comment, produced by osm2mp
                LastWayID = Trim(sLine)
            End If
            If SectionType = 1 Then
                'line of header section
                MPheader = MPheader + sLine + vbNewLine
            End If
        Case 1 'scan for routing param
            If Left(Trim(sLine), 10) = "RouteParam" Then
                If Left(Trim(sLine), 13) = "RouteParamExt" Then GoTo lNoData 'skip ext
                k2 = InStr(1, sLine, "=") + 1
                routep = Split(Mid(sLine, k2, Len(sLine) - k2), ",") 'split by "," delimiter
                If Control_ForceWaySpeed = -1 Then
                    WaySpeed = Val(routep(0))  'direct copy of speed
                Else
                    WaySpeed = Control_ForceWaySpeed
                End If
                WayOneway = Val(routep(2)) 'and oneway
                If LastCommentHighway = HIGHWAY_UNSPECIFIED Then
                    'WayClass = 3 'default class
                    If Control_LoadMPType = 0 Then WayClass = HIGHWAY_SECONDARY
                Else
                    'get class from osm2mp comment
                    WayClass = LastCommentHighway
                End If
            End If
            If Left(Trim(sLine), 5) = "Label" Then
                'label
                k2 = InStr(1, sLine, "=")
                label = Trim(Mid(sLine, k2 + 1, Len(sLine) - k2))
                If Left(label, 4) = "~[0x" Then
                    'use only special codes
                    k2 = InStr(1, label, "]")
                    label = Trim(Mid(label, k2 + 1, Len(label) - k2))
                Else
                    'ignore others
                    label = ""
                End If
                
            End If
            If Control_LoadMPType = 1 And Left(Trim(sLine), 5) = "Type=" Then
                'type
                k2 = InStr(1, sLine, "=0x")
                k3 = Val("&h" + Trim(Mid(sLine, k2 + 3, Len(sLine) - k2)))
                WayClass = GetTypeFromMP(k3)
                
            End If
        Case 3
            'scan for node routing info:
            If Left(Trim(sLine), 3) = "Nod" Then
                'Nod
                
                k2 = InStr(1, sLine, "=")
                If k2 <= 0 Then GoTo lSkipRoadNode
                k = Val(Mid(sLine, k2 + 1, 20))
                
                If k > NodesAlloc Then
                    'error: too big node index: " + sLine
                    GoTo lSkipRoadNode
                End If
                
                k3 = InStr(k2, sLine, ",")
                If k3 < 0 Then
                    'error: bad NodeID
                    GoTo lSkipRoadNode
                End If
                NodeID = Val(Mid(sLine, k3 + 1, 20))
                
                If NodeID > NodeIDMax Then NodeIDMax = NodeID 'update max NodeID
                
                Nodes(NodesNum - ThisLineNodes + k).NodeID = NodeID  'store nodeid
lSkipRoadNode:
            
            End If
            
        Case 2
        If Left(Trim(sLine), 4) = "Data" Then
            'geometry
            
            k = InStr(1, sLine, "Data")
            If k <= 0 Then GoTo lNoData
            
            sWay = ""
            sPrefix = Left(sLine, k + 4) + "=" '"Data" + next char + "="
            
            DataLineNum = DataLineNum + 1
            
            ThisLineNodes = 0
            LinkLen = 0
            ChainNum = 0
            p = NodesNum '1st node of way
            
lNextPoint:
            'get lat-lon coords from line
            k2 = InStr(k, sLine, "(")
            If k2 <= 0 Then GoTo lEndData
            fLat = Val(Mid(sLine, k2 + 1, 20))

            k3 = InStr(k2, sLine, ",")
            If k3 <= 0 Then GoTo lEndData
            fLon = Val(Mid(sLine, k3 + 1, 20))
            
lAddNode:
            'fill node info
            Nodes(NodesNum).lat = fLat
            Nodes(NodesNum).lon = fLon
            Nodes(NodesNum).Edges = 0
            Nodes(NodesNum).NodeID = -1
                        
            If ThisLineNodes > 0 Then
                'not the first node of way -> create edge
                j = JoinByEdge(NodesNum - 1, NodesNum)
                Edges(j).oneway = WayOneway 'oneway edges is always -> by geometry
                Edges(j).roadtype = WayClass
                Edges(j).label = label
                If WaySpeed >= 0 Then
                    Edges(j).speed = WaySpeed
                Else
                    'were not specified
                    Edges(j).speed = 3 '56km/h
                End If
                
                If (WayClass And HIGHWAY_MASK_LINK) > 0 Then
                    LinkLen = LinkLen + Distance(NodesNum - 1, NodesNum)
                    If LinkLen > MaxLinkLen Then
                        WayClass = WayClass And HIGHWAY_MASK_MAIN
                        Edges(j).roadtype = WayClass
                        For i = 0 To ChainNum - 1
                            Edges(Chain(i)).roadtype = WayClass
                        Next
                        'Debug.Print LastWayID; " - "; CStr(LinkLen) 'uncomment to log list of ways
                        NumDelinked = NumDelinked + 1
                    Else
                        Call AddChain(j)
                    End If
                End If
            End If
            ThisLineNodes = ThisLineNodes + 1
            
            Call AddNode 'finish node creation
            
            k = k3
            GoTo lNextPoint
            
lEndData:
        If SectionType = 3 Then
            'polygone - need to close it
            If p > -1 Then
                fLat = Nodes(p).lat
                fLon = Nodes(p).lon
                p = -1
                GoTo lAddNode
            End If
        End If
            
            
lNoData:
        End If
    End Select

    If Not EOF(1) Then GoTo lNextLine
    
    Close #1
    ChainNum = 0
    'Debug.Print "De-_link-ed: "; NumDelinked 'uncomment to log number of ways

End Sub



'Check that index is already present in Chain() array
Public Function IsInChain(node1 As Long) As Long
    Dim i As Long
    For i = 0 To ChainNum - 1
        If Chain(i) = node1 Then IsInChain = 1: Exit Function
    Next
    IsInChain = 0
End Function


'Find index of node1 in Chain() array, return -1 if not present
Public Function FindInChain(node1 As Long) As Long
    Dim i As Long
    For i = 0 To ChainNum - 1
        If Chain(i) = node1 Then FindInChain = i: Exit Function
    Next
    FindInChain = -1
End Function



'Check all edges of node for junction marker and add them into collapsing constuction
Public Sub CheckForCollapeByChain2(node1 As Long)
    
    Dim j As Long
    Dim edge As Long
    Dim BorderNode As Long
    
    BorderNode = 0
    
    For j = 0 To Nodes(node1).Edges - 1
        edge = Nodes(node1).edge(j)
        If (Edges(edge).mark And MARK_JUNCTION) > 0 Then
            'edge is marked as junction
            Edges(edge).mark = Edges(edge).mark Or MARK_COLLAPSING 'mark it as collapsing
            
            'add other end of edge into collapsing constuction
            If IsInChain(Edges(edge).node1) = 0 Then
                'node1 is not in chain
                Call AddChain(Edges(edge).node1)
            End If
            If IsInChain(Edges(edge).node2) = 0 Then
                'node2 is not in chain
                Call AddChain(Edges(edge).node2)
            End If
            GoTo lNext
        End If
        
        'at lease one non-junction edge found
        BorderNode = 1
lNext:
    Next
    
    If BorderNode = 1 Then Nodes(node1).mark = MARK_NODE_BORDER  'node is border-node

End Sub


'Check all edges of node for collapsing marker and add their other ends into chain
Public Sub GroupCollapse(node1 As Long)
    Dim j As Long
    Dim edge As Long
    Dim k As Long
    
    For j = 0 To Nodes(node1).Edges - 1
        edge = Nodes(node1).edge(j)
        If (Edges(edge).mark And MARK_COLLAPSING) > 0 Then
            k = Edges(edge).node1
            If k = node1 Then k = Edges(edge).node2 'get other end of edge
            
            'if other end is marked as node-of-junction -> add it to current chain
            If Nodes(k).mark = MARK_NODE_OF_JUNCTION And IsInChain(k) = 0 Then Call AddChain(k)
        End If
    Next
End Sub

'Collapse all junctions to nodes
'junctions detected as clouds of _links and short loops
'SlideMaxDist - max distance allowed to slide by border-nodes
'LoopLimit - max loop considered as junction
'AngleLimit - min angle between aiming edges to use precise calc of centroid
Public Sub CollapseJunctions2(SlideMaxDist As Double, LoopLimit As Double, AngleLimit As Double)
    Dim i As Long, j As Long, k As Long, mode1 As Long
    Dim m As Long, e As Long
    Dim JoinGroups As Long
    Dim JoinedNode() As Long 'controids
    Dim JoiningNodes As Long
    Dim PassNumber As Long
    Dim BorderNodes As Long
    Dim JoinGroupBox() As bbox
    
    PassNumber = 1
    
    'Algorithm marks _link edges as junctions
    'then all edges checked for participating of short loops, if short loop found all its edges also marked as junctions
    'then all junction edges tries to collapse with keeping connection points (border nodes) on other road
    'border nodes can slide on other edges while construction of junction edges tries to shrink like stretched rubber
    'siding also marks passed edges as part of collapsing junctions
    'final constructions from collapsing edges is separated from each other
    'then centroids of this constructions found, it is done by finding point which minimizes sum of squares
    'of distances to aiming lines, aiming lines is build from all edges connecting to collapsing junction and
    'all not _link edges of this junctions
    'then all collapsing edges deleted, all nodes - joined into centroid
    'then algorithm reiterates from start till no junction were found
    
lIteration:

    'mark all as not-yet-joining
    For i = 0 To NodesNum - 1
        Nodes(i).mark = 0 'no wave
        Nodes(i).temp_dist = 0 'no wave
    Next
    
    '1) Mark all potential parts of junctions
    For i = 0 To EdgesNum - 1
        Edges(i).mark = 0 'not checked
        If Edges(i).node1 > -1 Then
            If (Edges(i).roadtype And HIGHWAY_MASK_LINK) <> 0 Then
                'all links are parts of junctions
                Edges(i).mark = MARK_JUNCTION
            End If
        End If
    Next
    
    For i = 0 To EdgesNum - 1
        If Edges(i).node1 > -1 Then
            'check all edges (not links) for short loops
            'all short loops also marked as part of junctions
            
            If (Edges(i).roadtype And HIGHWAY_MASK_LINK) = 0 And (Edges(i).mark And MARK_JUNCTION) = 0 Then
                'not _link, not yet marked junction
                If (Edges(i).mark And MARK_DISTCHECK) = 0 Then Call CheckShortLoop2(i, LoopLimit) 'not marked distcheck -> should check it
            End If
        End If
        If (i And 16383) = 0 Then
            'display progress
            Form1.Caption = "CollapseJunctions2 (" + CStr(SlideMaxDist) + ", " + CStr(LoopLimit) + ", " + CStr(AngleLimit) + ") #" + CStr(PassNumber) + ", Shorts " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    Next
    
    For i = 0 To NodesNum - 1
        Nodes(i).mark = -1 'not in join group
    Next

    '2) mark edges by trying to shrink junction
    JoinGroups = 0
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED And Nodes(i).mark = -1 Then
            'not deleted node, not marked yet -> should check
             
            'start new group
            ChainNum = 0
            BorderNodes = 0
            Call AddChain(i) 'add this node to a chain
            
            Call CheckForCollapeByChain2(Chain(0)) 'check all edges of node for collapsing
            
            If Nodes(Chain(0)).mark = MARK_NODE_BORDER Then BorderNodes = 1  'node is border-node
            
            j = 1
            If ChainNum > 1 Then
                'at least 2 nodes found to collapse
                
lRecheckAgain:
                'continue to check all edges of added nodes by chain
                While j < ChainNum
                    
                    Call CheckForCollapeByChain2(Chain(j))
                    
                    If Nodes(Chain(j)).mark = MARK_NODE_BORDER Then
                        'border-node found - move it to start of chain (by swap)
                        k = Chain(BorderNodes)
                        Chain(BorderNodes) = Chain(j)
                        Chain(j) = k
                        
                        BorderNodes = BorderNodes + 1
                    End If
                    j = j + 1
                Wend
                
                If BorderNodes > 1 Then
                    'if border-nodes found - shrink whole construction by sliding border-nodes by geometry
                    Call ShrinkBorderNodes(BorderNodes, SlideMaxDist)
                    BorderNodes = 0
                    GoTo lRecheckAgain
                End If
                
                'mark all found nodes
                For j = 0 To ChainNum - 1
                    Nodes(Chain(j)).mark = MARK_NODE_OF_JUNCTION
                Next
            End If
        End If
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "CollapseJunctions2 (" + CStr(SlideMaxDist) + ", " + CStr(LoopLimit) + ", " + CStr(AngleLimit) + ") #" + CStr(PassNumber) + ", Shrink " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next
    
    '3) group edges to separate junctions
    JoinGroups = 0
    For i = 0 To NodesNum - 1
        If Nodes(i).mark = MARK_NODE_OF_JUNCTION Then
            ChainNum = 0
            Call AddChain(i) 'add this node to a chain
            j = 0
            While j < ChainNum
                Call GroupCollapse(Chain(j)) 'check all edges and add their other ends if they are collapsing
                j = j + 1
            Wend
            
            If ChainNum > 1 Then
                '2 or more found - new group
                For j = 0 To ChainNum - 1
                    Nodes(Chain(j)).mark = JoinGroups
                Next
                JoinGroups = JoinGroups + 1
            End If
        End If
    Next
    
    '4) calculate coordinates of centroid for collapsed junction
    
    ReDim JoinedNode(JoinGroups)
    ReDim JoinGroupBox(JoinGroups)
    
    'Create nodes for all found join-groups
    For i = 0 To JoinGroups
        JoinedNode(i) = NodesNum
        Nodes(NodesNum).mark = -1
        Nodes(NodesNum).Edges = 0
        Nodes(NodesNum).lat = Nodes(0).lat 'fake coords, so cluster-index algo will not get (0,0) coords
        Nodes(NodesNum).lon = Nodes(0).lon
        Call AddNode
        JoinGroupBox(i).lat_max = -360
        JoinGroupBox(i).lat_min = 360
        JoinGroupBox(i).lon_max = -360
        JoinGroupBox(i).lon_min = 360
    Next
    
    'calc bboxes for all join groups
    For j = 0 To NodesNum - 1
        If Nodes(j).mark >= 0 Then
            i = Nodes(j).mark
            
            'update bbox for group
            If Nodes(j).lat < JoinGroupBox(i).lat_min Then JoinGroupBox(i).lat_min = Nodes(j).lat
            If Nodes(j).lat > JoinGroupBox(i).lat_max Then JoinGroupBox(i).lat_max = Nodes(j).lat
            If Nodes(j).lon < JoinGroupBox(i).lon_min Then JoinGroupBox(i).lon_min = Nodes(j).lon
            If Nodes(j).lon > JoinGroupBox(i).lon_max Then JoinGroupBox(i).lon_max = Nodes(j).lon
        End If
    Next
    
    Call BuildNodeClusterIndex(0) 'rebuild cluster-index from zero
    
    For i = 0 To JoinGroups - 1
        JoiningNodes = 0
        AimEdgesNum = 0
        
        mode1 = 0 'first
lNextNode:
        'get node from bbox by cluster-index
        'without cluster-index search have complexety ~O(n^2)
        j = GetNodeInBboxByCluster(JoinGroupBox(i), mode1)
        mode1 = 1 '"next" next time
        
        If j <> -1 Then
            If Nodes(j).mark = i Then
                'this joining group
                JoiningNodes = JoiningNodes + 1
                For m = 0 To Nodes(j).Edges - 1
                    e = Nodes(j).edge(m)
                    If (Edges(e).mark And MARK_AIMING) > 0 Then GoTo lSkipEdge 'skip already marked as aiming
                    If (Edges(e).mark And MARK_COLLAPSING) > 0 And (Edges(e).mark And MARK_JUNCTION) > 0 Then GoTo lSkipEdge 'skip all collapsing junctions
                    
                    'remain: all edges which will survive from all nodes of this join group
                    'plus all collapsing edges of main roads (this needed for removing noise of last edges near junctions)
                    If Edges(e).node1 = j Then
                        AimEdges(AimEdgesNum).lat1 = Nodes(j).lat 'lat1-lon1 is always a node which will collapse (and so points to junction)
                        AimEdges(AimEdgesNum).lon1 = Nodes(j).lon
                        AimEdges(AimEdgesNum).lat2 = Nodes(Edges(e).node2).lat
                        AimEdges(AimEdgesNum).lon2 = Nodes(Edges(e).node2).lon
                    Else
                        AimEdges(AimEdgesNum).lat1 = Nodes(j).lat
                        AimEdges(AimEdgesNum).lon1 = Nodes(j).lon
                        AimEdges(AimEdgesNum).lat2 = Nodes(Edges(e).node1).lat
                        AimEdges(AimEdgesNum).lon2 = Nodes(Edges(e).node1).lon
                    End If
                    If (Edges(Nodes(j).edge(m)).mark And MARK_COLLAPSING) > 0 Then
                        'mark as aiming, for skip it next time
                        Edges(e).mark = Edges(e).mark Or MARK_AIMING
                    End If
                    Call AddAimEdge
lSkipEdge:
                Next
            End If
            GoTo lNextNode
        End If
        
        If JoiningNodes > 0 Then
            If (JoiningNodes And 127) = 0 Then
                'display progress
                Form1.Caption = "CollapseJunctions2 (" + CStr(SlideMaxDist) + ", " + CStr(LoopLimit) + ", " + CStr(AngleLimit) + ") #" + CStr(PassNumber) + ", Aim " + CStr(i) + " / " + CStr(JoinGroups): Form1.Refresh
            End If
            Call FindAiming(Nodes(JoinedNode(i)), AngleLimit) 'find centroid of junction
        End If
    Next
    
    '5) delete all collapsing edges
    For i = 0 To EdgesNum - 1
        If (Edges(i).mark And MARK_COLLAPSING) > 0 Then
            Call DelEdge(i)
        End If
    Next
    
    '6) Ñollapse nodes to junctions single nodes
    For j = 0 To NodesNum - 1
        If Nodes(j).mark >= 0 Then
            i = Nodes(j).mark
            For m = 0 To Nodes(j).Edges - 1
                'reconnect edge to centroid
                If Edges(Nodes(j).edge(m)).node1 = j Then
                    Edges(Nodes(j).edge(m)).node1 = JoinedNode(i)
                Else
                    Edges(Nodes(j).edge(m)).node2 = JoinedNode(i)
                End If
                Call AddEdgeToNode(JoinedNode(i), Nodes(j).edge(m))
                'k = Nodes(JoinedNode(i)).Edges
                'Nodes(JoinedNode(i)).edge(k) = Nodes(j).edge(m)
                'Nodes(JoinedNode(i)).Edges = k + 1
            Next
            Nodes(j).Edges = 0 'all edges were reconnected
            Call DelNode(j) 'kill
        End If
        If (j And 8191) = 0 Then
            Form1.Caption = "CollapseJunctions2 (" + CStr(SlideMaxDist) + ", " + CStr(LoopLimit) + ", " + CStr(AngleLimit) + ") #" + CStr(PassNumber) + ", Del " + CStr(j) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next
    
    If JoinGroups > 0 Then
        'unless no junction detected - relaunch algo
        PassNumber = PassNumber + 1
        DoEvents
        GoTo lIteration
    End If

End Sub


'Shrink group of border-nodes to minimum-distance (point or segment or more complex)
'BorderNum - numbed of border nodes (must be in the start of Chain array)
'shrink is achived by moving border-nodes along geometry while minimizing sum length
'all edges, passed by border-node marked as part of junctions
'MaxShift limit max length allowed to pass by each border-node
'if two border-nodes reach each other, they joins
'if reaching 1 border-node is not possible, then near edges are checked for internal-points minimizing sum len
'this edges also marked as part of junctions (requires than all edges were not very long)
Public Sub ShrinkBorderNodes(ByVal BorderNum As Long, MaxShift As Double)
    Dim BorderNodes() As Long 'current node index of border-node
    Dim BorderShifts() As Double 'distance, covered by border-node while moving on edges
    Dim i As Long, j As Long, k As Long
    Dim e As Long, p As Long
    Dim Moving As Long
    Dim Moved As Long
    Dim Dist As Double, dist0 As Double
    Dim dist1 As Double
    Dim dist_min As Double, node_dist_min As Long, edge_dist_min As Long
    ReDim BorderNodes(BorderNum) 'indexes of border-nodes
    ReDim BorderShifts(BorderNum) 'len, passed by border-nodes
    
    'get border nodes
    For i = 0 To BorderNum - 1
        BorderNodes(i) = Chain(i)
        BorderShifts(i) = 0 'zero len passed
    Next
    
lRestartCycle:
    Moving = 0
    Moved = 0
    
    While Moving < BorderNum
lRestartMoving:
        'calc current sum of distances from Moving border-node to all others
        dist0 = 0
        For j = 0 To BorderNum - 1
            If j <> Moving Then
                dist1 = Distance(BorderNodes(Moving), BorderNodes(j))
                dist0 = dist0 + dist1
            End If
        Next
                
        'finding node, where this border-node can move to minimize distance
        dist_min = dist0 'need sum-distance less than current
        node_dist_min = -1: edge_dist_min = -1 'not yet found
        For k = 0 To Nodes(BorderNodes(Moving)).Edges - 1
            e = Nodes(BorderNodes(Moving)).edge(k)
            If (Edges(e).mark And MARK_JUNCTION) = 0 Then
                'not junction edge
                p = Edges(e).node1
                dist1 = Distance(p, Edges(e).node2) 'get len of this edge
                If BorderShifts(Moving) + dist1 > MaxShift Then GoTo lSkipAsLong 'moving will exceed MaxShift
                If p = BorderNodes(Moving) Then p = Edges(e).node2 'get other end of edge
                
                'calc new sum of distances from this border-node to all others
                Dist = 0
                For j = 0 To BorderNum - 1
                    If j <> Moving Then
                        dist1 = Distance(p, BorderNodes(j))
                        Dist = Dist + dist1
                    End If
                Next
                If Dist < dist_min Then dist_min = Dist: node_dist_min = p: edge_dist_min = e 'minimizing found
lSkipAsLong:
            End If
        Next
        
        If node_dist_min > -1 Then
            'found node, more close to other border-nodes
            Edges(edge_dist_min).mark = Edges(edge_dist_min).mark Or MARK_COLLAPSING   'mark for collapse
            If IsInChain(node_dist_min) = 0 Then Call AddChain(node_dist_min)  'add found node to chain of junction nodes
            Moved = 1 'at least 1 border-node moved
            
            BorderShifts(Moving) = BorderShifts(Moving) + Distance(Edges(edge_dist_min).node1, Edges(edge_dist_min).node2) ' update passed distance
            
            For i = 0 To BorderNum - 1
                If node_dist_min = BorderNodes(i) Then
                    'after moving border-node joined with another one - remove
                    If BorderShifts(i) < BorderShifts(Moving) Then
                        'joined with node with smaller moves - keep smallest
                        BorderShifts(i) = BorderShifts(Moving)
                    End If
                    
                    'join
                    BorderNodes(Moving) = BorderNodes(BorderNum - 1)
                    BorderShifts(Moving) = BorderShifts(BorderNum - 1)
                    BorderNum = BorderNum - 1
                    If BorderNum = 1 Then GoTo lReachOne 'only 1 border-node left
                    GoTo lRestartMoving 'back to moving this node
                End If
            Next
            
            'border-node not joined, just moved
            BorderNodes(Moving) = node_dist_min
            GoTo lRestartMoving
        End If
        'not found - proceeding to next node
        
        Moving = Moving + 1
    Wend
    
    If Moved = 1 Then GoTo lRestartCycle 'some border-nodes were moved, repeat cycle
    
    'no border-nodes moved during whole cycle - looks like shrinking reached minimum
    
    If BorderNum > 1 Then
        '2 or more border nodes remains
        
        For i = 0 To BorderNum - 1 'all border-nodes
            For k = 0 To Nodes(BorderNodes(i)).Edges - 1 'all edges of it
                e = Nodes(BorderNodes(i)).edge(k)
                If (Edges(e).mark And MARK_JUNCTION) = 0 Then
                    'not junction edge -> edge of main road
                    For j = 0 To BorderNum - 1
                        If j <> i Then 'not the same border-node
                            Dist = DistanceToSegment(Edges(e).node1, Edges(e).node2, BorderNodes(j))
                            If DistanceToSegment_last_case = 3 Then
                                '3rd case distance - interval internal point is closer to border-node j than both ends of interval
                                '-> mark edge and both nodes for collapsing
                                Edges(e).mark = Edges(e).mark Or MARK_COLLAPSING  'mark for definite collapse
                                If IsInChain(Edges(e).node1) = 0 Then Call AddChain(Edges(e).node1)
                                If IsInChain(Edges(e).node2) = 0 Then Call AddChain(Edges(e).node2)
                            End If
                        End If
                    Next
                End If
            Next
        Next
    End If
    
lReachOne:
    'one border-node reach - end
    
End Sub


'mark one half of loop as junction by moving from node to node with smallest temp_dist
'node1 - start node (i.e. where waves collide)
Public Sub MarkLoopHalf(node1 As Long)
    Dim i As Long, e As Long, d As Long
    Dim min_dist As Double, min_dist_node As Long, min_dist_edge As Long
    Dim node As Long
    
    node = node1
    
lMoveNext:
    If Nodes(node).mark = 1 Or Nodes(node).mark = -1 Then Exit Sub 'reached start node - exit
    
    'find near node with smaller temp_dist
    min_dist = Nodes(node).temp_dist
    min_dist_node = -1
    min_dist_edge = -1
    For i = 0 To Nodes(node).Edges - 1 'check all edges
        e = Nodes(node).edge(i)
        d = Edges(e).node1
        If d = node Then d = Edges(e).node2
        'd - always other end of edge
        If Nodes(d).mark <> 0 And Nodes(d).temp_dist < min_dist Then
            'smaller temp_dist found
            min_dist = Nodes(d).temp_dist
            min_dist_node = d
            min_dist_edge = e
        End If
    Next
    
    If min_dist_edge > -1 Then
        Edges(min_dist_edge).mark = Edges(min_dist_edge).mark Or (MARK_JUNCTION) 'mark passed edge as junction
        node = min_dist_node
        GoTo lMoveNext
    End If

End Sub


'Check edge for participating in short loop (shorted than MaxDist)
'Launch two waves for propagation - one from each end of edge
'if waves collide, they are part of short loop
'waves are limited by length and MARK_DISTCHECK
'if no short loop found, this edge marked with MARK_DISTCHECK (means, no short loop passing this edge)
Public Sub CheckShortLoop2(edge1 As Long, MaxDist As Double)
    Dim i As Long, j As Long, k As Long, k2 As Long
    Dim e As Long, d As Long, q As Long
    Dim node1 As Long
    Dim node2 As Long
    Dim dist0 As Double, dist1 As Double
    
    'wave starts
    node1 = Edges(edge1).node1
    node2 = Edges(edge1).node2
    dist0 = 0.5 * Distance(node1, node2) 'half of edge len - start point is center of edge
    Nodes(node1).mark = 1
    Nodes(node2).mark = -1
    Edges(edge1).mark = Edges(edge1).mark Or MARK_WAVEPASSED
    ChainNum = 0
    Call AddChain(node1)
    Call AddChain(node2)
    Nodes(node1).temp_dist = dist0
    Nodes(node2).temp_dist = dist0
    
    'propagate waves
    j = 0
    While j < ChainNum
        k = Nodes(Chain(j)).mark
        If k > 0 Then
            k2 = k + 1
        Else
            k2 = k - 1
        End If
        
        dist0 = Nodes(Chain(j)).temp_dist
        For i = 0 To Nodes(Chain(j)).Edges - 1
            q = Nodes(Chain(j)).edge(i)
            If (Edges(q).mark And MARK_WAVEPASSED) <> 0 Then GoTo lSkipEdge 'wave already passed this edge
            If (Edges(q).mark And MARK_DISTCHECK) <> 0 Then GoTo lSkipEdge 'no short loop here - no need to pass thru this edge
            Edges(q).mark = Edges(q).mark Or MARK_WAVEPASSED
            d = Edges(q).node1
            If d = Chain(j) Then d = Edges(q).node2
            If Nodes(d).mark <> 0 Then
                If (Nodes(d).mark < 0 And k > 0) Or (Nodes(d).mark > 0 And k < 0) Then
                    'loop found
                    dist1 = dist0 + Distance(d, Chain(j)) 'update by len of this edge
                    dist1 = dist1 + Nodes(d).temp_dist 'len of second part of wave
                    If dist1 > MaxDist Then GoTo lSkipEdge 'loop is too long
                    
                    GoTo lShortLoop 'short loop found
                End If
                GoTo lSkipEdge
            End If
            dist1 = dist0 + Distance(d, Chain(j)) 'update by len of this edge
            Nodes(d).temp_dist = dist1 'set passed len
            Nodes(d).mark = k2
            If dist1 < MaxDist Then Call AddChain(d) 'add to chain, but only if distance from start is not too long
lSkipEdge:
        Next
        j = j + 1
    Wend
    'short loop not found
    
    Edges(edge1).mark = Edges(edge1).mark Or MARK_DISTCHECK 'no short passing this edge
    GoTo lClearTemp
    
lShortLoop:
    'short loop found
    Edges(q).mark = Edges(q).mark Or (MARK_JUNCTION) 'mark final edge
    
    'mark both loop half by moving backward from collision edge to start
    Call MarkLoopHalf(d)
    Call MarkLoopHalf(Chain(j))
    
    Edges(edge1).mark = Edges(edge1).mark Or (MARK_JUNCTION) 'mark start edge
    
lClearTemp:
    'clear all temp marks
    For j = 0 To ChainNum - 1
        Nodes(Chain(j)).mark = 0
        Nodes(Chain(j)).temp_dist = 0
        For i = 0 To Nodes(Chain(j)).Edges - 1
            q = Nodes(Chain(j)).edge(i)
            Edges(q).mark = Edges(q).mark And (-1 Xor MARK_WAVEPASSED)
        Next
    Next
    
End Sub

'Find centroid of junction by aiming-edges
'Will find location, equally distant from all lines (defined by AimEdges)
'Iterative search by 5 points, combines moving into direction of minimizing sum of squares of distances
'and bisection method to clarify centroid position
Public Sub FindAiming(ByRef result As node, AngleLimit As Double)
    Dim px As Double
    Dim py As Double
    Dim dx As Double
    Dim dy As Double
    Dim q As Double
    Dim i As Long, j As Long, t As Long
    Dim v As Double, v1 As Double, v2 As Double, v3 As Double, v4 As Double
    Dim dvx As Double, dvy As Double
    Dim EstStepX As Double
    Dim EstStepY As Double
    'Dim phase As Long
    Dim MaxAngle As Double
    
    px = 0
    py = 0
    
    If AimEdgesNum = 0 Then Exit Sub
    
    'calculate equation elements of all aimedges
    'equation: Distance of (x,y) = a * x + b * y + c
    'q[i] = 1 / sqrt(((y2[i]-y1[i])^2+(x2[i]-x1[i])^2)
    'a[i] = (y2[i]-y1[i])*q[i]
    'b[i] = (x1[i]-x2[i])*q[i]
    'c[i] = (y1[i]*x2[i]-x1[i]*y2[i])*q[i]
    'also calc default centroid as average of all aimedges lat1-lot1 coords
    
    For i = 0 To AimEdgesNum - 1
        'TODO: fix (not safe to 180/-180 edge)
        px = px + AimEdges(i).lat1
        py = py + AimEdges(i).lon1
        dx = AimEdges(i).lat2 - AimEdges(i).lat1
        dy = AimEdges(i).lon2 - AimEdges(i).lon1
        q = Sqr(dx * dx + dy * dy)
        AimEdges(i).d = q
        If q <> 0 Then
            q = 1 / q
        End If
        AimEdges(i).a = dy * q ' a and b is normalized normal to edge
        AimEdges(i).b = -dx * q
        AimEdges(i).c = (AimEdges(i).lon1 * AimEdges(i).lat2 - AimEdges(i).lon2 * AimEdges(i).lat1) * q
    Next
    px = px / AimEdgesNum
    py = py / AimEdgesNum
    
    'check max angle between aimedges
    'angle is checked in lat/lon grid, so not exactly angle in real world
    MaxAngle = 0
    For i = 1 To AimEdgesNum - 1
        For j = 0 To i - 1
            'vector product of normals, result is sin(angle)
            'angle is smallest of (a,180-a)
            q = Abs(AimEdges(i).a * AimEdges(j).b - AimEdges(i).b * AimEdges(j).a)
            If q > MaxAngle Then MaxAngle = q
        Next
    Next
    
    If MaxAngle < AngleLimit Then GoTo lResult 'angle is too small, iterative aiming will make big error along roads and should not be used
    
    
    'OK, angle is good => lets start iterative search
    
    'initial steps
    EstStepX = 0.0001
    EstStepY = 0.0001
    t = 0
    'px and py is start location
    
lNextStep:
    t = t + 1
    v = 0
    v1 = 0
    v2 = 0
    v3 = 0
    v4 = 0
    
    'calc distance in 5 points - current guess and 4 points on ends of "+"-cross
    For i = 0 To AimEdgesNum - 1
        'sum of module distances, not good
'        v = v + Abs(px * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c)
'        v1 = v1 + Abs((px + EstStepX) * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c)
'        v2 = v2 + Abs((px - EstStepX) * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c)
'        v3 = v3 + Abs(px * AimEdges(i).a + (py + EstStepY) * AimEdges(i).b + AimEdges(i).c)
'        v4 = v4 + Abs(px * AimEdges(i).a + (py - EstStepY) * AimEdges(i).b + AimEdges(i).c)

        'sum of square distances - better
        v = v + (px * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c) ^ 2
        v1 = v1 + ((px + EstStepX) * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c) ^ 2
        v2 = v2 + ((px - EstStepX) * AimEdges(i).a + py * AimEdges(i).b + AimEdges(i).c) ^ 2
        v3 = v3 + (px * AimEdges(i).a + (py + EstStepY) * AimEdges(i).b + AimEdges(i).c) ^ 2
        v4 = v4 + (px * AimEdges(i).a + (py - EstStepY) * AimEdges(i).b + AimEdges(i).c) ^ 2
    Next
    
    If v > v1 Or v > v2 Or v > v3 Or v > v4 Then
        'v is not smallest => centroid location is not in covered by our cross (px+-EstStepX,py+-EstStepY)
        '=> we need to shift
        If v > v1 Or v > v2 Then
            'shift by X (by half of quad) in direction to minimize v
            If v1 < v2 Then px = px + EstStepX * 0.5 Else px = px - EstStepX * 0.5
        End If
        If v > v3 Or v > v4 Then
            'shift by Y (by half of quad) in direction to minimize v
            If v3 < v4 Then py = py + EstStepY * 0.5 Else py = py - EstStepY * 0.5
        End If
        GoTo lNextStep
    Else
        'v is smallest => centroid location IS covered by our cross (px+-EstStepX,py+-EstStepY)
        'we need to select sub-rectangle to clarify position
        
        'find q as max of v1-v4
        q = v1: i = 1
        If v2 > q Then q = v2: i = 2
        If v3 > q Then q = v3: i = 3
        If v4 > q Then q = v4: i = 4
        Select Case i
            Case 4
                'v4 is max, select half with v3
                py = py + EstStepY * 0.5
                EstStepY = EstStepY * 0.5
            Case 3
                'v3 is max, select half with v4
                py = py - EstStepY * 0.5
                EstStepY = EstStepY * 0.5
            Case 2
                'v2 is max, select half with v1
                px = px + EstStepX * 0.5
                EstStepX = EstStepX * 0.5
            Case 1
                'v1 is max, select half with v2
                px = px - EstStepX * 0.5
                EstStepX = EstStepX * 0.5
        End Select
        
        'if required accuracy not yet reached - continue
        'exit if 100k iteration does not help
        If t < 100000 And EstStepX > 0.0000001 Or EstStepY > 0.0000001 Then GoTo lNextStep
    End If
    
    'OK, found
    
lResult:
    result.lat = px
    result.lon = py
    
End Sub

'Calc speedclass and label of combined subchain of edges
Public Sub EstimateChain(IndexStart As Long, IndexLast As Long)
    Dim i As Long, j As Long
    For i = 0 To 10
        SpeedHistogram(i) = 0
    Next
    
    EstimateChain_label = ""
    EstimateChain_speed = 0
    ResetLabelStats
    
    For i = IndexStart To IndexLast - 1
        j = GetEdgeBetween(Chain(i), Chain(i + 1))
        If j >= 0 Then
            Call AddLabelStat0(Edges(j).label) 'add label of edge into stats
            j = Edges(j).speed
            SpeedHistogram(j) = SpeedHistogram(j) + 1 'add speed into histogram
        End If
    Next
    
    EstimateChain_speed = EstimateSpeedByHistogram 'estimate speed
    EstimateChain_label = GetLabelByStats(0) 'calc resulting label
    
End Sub


'Estimate speed class of road by histogram
Public Function EstimateSpeedByHistogram() As Integer
    Dim total As Long, i As Long
    
    'call total sum
    For i = 0 To 10
        total = total + SpeedHistogram(i)
    Next

    EstimateSpeedByHistogram = 3 'default speedclass
    
    If total = 0 Then Exit Function 'should never happens
    
    'find speedclass with 90% coverage
    For i = 0 To 10
        If SpeedHistogram(i) > total * 0.9 Then
            '90% of chain have this speedclass
            EstimateSpeedByHistogram = i
            Exit Function
        End If
    Next
    
    'no 90%
    
    'find minimum speedclass with 40% coverage
    For i = 0 To 10
        If SpeedHistogram(i) > total * 0.4 Then
            '40% of chain have this speedclass
            EstimateSpeedByHistogram = i
            Exit Function
        End If
    Next
    
    'no 40% (very much alike will not happens)
    
    'find minimum speedclass with 10% coverage
    For i = 0 To 10
        If SpeedHistogram(i) > total * 0.1 Then
            '10% of chain have this speedclass
            EstimateSpeedByHistogram = i
            Exit Function
        End If
    Next
    
    'no 10% (almost impossible)
    
    'find minimum speedclass with >0 coverage
    For i = 0 To 10
        If SpeedHistogram(i) > 0 Then
            EstimateSpeedByHistogram = i
            Exit Function
        End If
    Next

End Function


'Delete edges which connect node with itself
Public Sub FilterVoidEdges()
    Dim i As Long
    For i = 0 To EdgesNum - 1
        If Edges(i).node1 <> -1 And Edges(i).node1 = Edges(i).node2 Then
            Call DelEdge(i)
        End If
    Next
End Sub


'Combine edges. edge2 is deleted, edge1 is kept
'assumed, that edges have at leaset 1 common node
'return: 0 - not combined, 1 - combined
Public Function CombineEdges(edge1 As Long, edge2 As Long, Optional CommonNode As Long = -1)
    CombineEdges = 0
    If CombineEdgeParams(edge1, edge2, CommonNode) > 0 Then
        CombineEdges = 1
        Call DelEdge(edge2)
    End If
End Function


'Compare roadtype-s
'return: 1 - type1 have higher priority, -1 - type2 have higher priority, 0 - equal
Public Function CompareRoadtype(ByVal type1 As Long, ByVal type2 As Long) As Long
    If type1 = type2 Then
        CompareRoadtype = 0 'just equal
    ElseIf (type1 And HIGHWAY_MASK_LINK) <> 0 And (type2 And HIGHWAY_MASK_LINK) = 0 Then
        CompareRoadtype = -1 'type1 is link, type2 is not
    ElseIf (type1 And HIGHWAY_MASK_LINK) = 0 And (type2 And HIGHWAY_MASK_LINK) <> 0 Then
        CompareRoadtype = 1 'type2 is link, type1 is not
    ElseIf (type1 And HIGHWAY_MASK_MAIN) < (type2 And HIGHWAY_MASK_MAIN) Then
        CompareRoadtype = 1 'type1 is less numerically - higher
    ElseIf (type1 And HIGHWAY_MASK_MAIN) > (type2 And HIGHWAY_MASK_MAIN) Then
        CompareRoadtype = -1 'type1 is higher numerically - less
    Else
        'should not happen
        CompareRoadtype = 0
    End If
End Function


'Combine edge parameters and store it to edge1
'return: 0 - not possible to combine, 1 - combined
Public Function CombineEdgeParams(edge1 As Long, edge2 As Long, Optional ByVal CommonNode As Long = -1)
    Dim k1 As Long
    Dim k2 As Long
    Dim k3 As Long
    CombineEdgeParams = 0
    
    If CommonNode = -1 Then
        'common node not specified in call
        If Edges(edge1).node1 = Edges(edge2).node1 Or Edges(edge1).node1 = Edges(edge2).node2 Then
            CommonNode = Edges(edge1).node1
        ElseIf Edges(edge1).node2 = Edges(edge2).node1 Or Edges(edge1).node2 = Edges(edge2).node2 Then
            CommonNode = Edges(edge1).node2
        Else
            'can't combine edges without at least one common point
            Exit Function
        End If
    End If
    
    'calc combined label - by stats
    Call ResetLabelStats
    Call AddLabelStat0(Edges(edge1).label)
    Call AddLabelStat0(Edges(edge2).label)
    Edges(edge1).label = GetLabelByStats(0)
    
    'calc combiner road type
    If Edges(edge1).roadtype <> Edges(edge2).roadtype Then
        'combine main road type - higher by OSM
        k1 = Edges(edge1).roadtype And HIGHWAY_MASK_MAIN
        k2 = Edges(edge2).roadtype And HIGHWAY_MASK_MAIN
        k3 = (Edges(edge2).roadtype And Edges(edge2).roadtype And HIGHWAY_MASK_LINK) 'keep link only if both are links
        If k2 < k1 Then k1 = k2 'numerically min roadtype
        Edges(edge1).roadtype = k1 Or k3
    End If
    
    'combined speed - lower
    If Edges(edge2).speed < Edges(edge1).speed Then
        
        Edges(edge1).speed = Edges(edge2).speed
    End If
    
    If Edges(edge1).oneway = 1 Then
        'combined oneway - keep only if both edges directed in one way
        If Edges(edge2).oneway = 1 Then
            'both edges are oneway
            If Edges(edge1).node1 = CommonNode And Edges(edge2).node2 = CommonNode Then
                'edges are opposite oneway, result is bidirectional
                Edges(edge1).oneway = 0
            ElseIf Edges(edge1).node2 = CommonNode And Edges(edge2).node1 = CommonNode Then
                'edges are opposite oneway, result is bidirectional
                Edges(edge1).oneway = 0
            End If
            
            'else - result is oneway
        Else
            'edge2 is bidirectional, so result also
            Edges(edge1).oneway = 0
        End If
    End If
    'if edge1 is bidirectional, so also result
    
    CombineEdgeParams = 1
End Function


'Join edges with very acute angle into one
'1) distance between edges ends < JoinDistance
'2) angle between edges lesser than limit
'AcuteKoeff: 1/tan() of limit angle  (3 =>18.4 degrees)
Public Sub JoinAcute(JoinDistance As Double, AcuteKoeff As Double)
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long
    Dim q As Long
    Dim Dist As Double
    Dim Merged As Long
    Dim PassNumber As Long
    
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED And Nodes(i).Edges > 1 Then
            Nodes(i).mark = 1 'mark to check, not deleted with 2+ edges
        Else
            Nodes(i).mark = 0 'mark to skip
        End If
    Next
    
    PassNumber = 1
    
lIteration:
    Merged = 0
    For i = 0 To NodesNum - 1
        
        If Nodes(i).mark = 1 Then
        
            'check for edges connecting same nodes several times
            'made by filling Chain array with other ends of edges
            ChainNum = 0
            j = 0
            While j < Nodes(i).Edges
                k = Edges(Nodes(i).edge(j)).node1
                If k = i Then k = Edges(Nodes(i).edge(j)).node2 'get other end
                m = FindInChain(k)
                If m = -1 Then
                    'first occurence in chain
                    Call AddChain(k)
                Else
                    'not first - should join
                    m = GetEdgeBetween(i, k)
                    If CombineEdges(m, Nodes(i).edge(j), i) > 0 Then GoTo lAgain 'combining succeed, we should check j-th edge once again
                End If
                j = j + 1
lAgain:
            Wend
            
            Nodes(i).mark = 0 'node is processed, mark to skip
            
            For j = 0 To ChainNum - 1
                If Chain(j) = -1 Then GoTo lSkipJ 'skip removed nodes
                If Nodes(Chain(j)).NodeID = MARK_NODEID_DELETED Then GoTo lSkipJ 'skip deleted nodes
                For k = 0 To ChainNum - 1
                    If j = k Or Chain(k) = -1 Then GoTo lSkipK 'skip same and removed nodes
                    If Nodes(Chain(k)).NodeID = MARK_NODEID_DELETED Then GoTo lSkipK 'skip deleted nodes
                    Dist = DistanceToSegment(i, Chain(j), Chain(k)) 'distance from Chain(k) to interval i-Chain(j)
                    If Dist < JoinDistance Then
                        'node Chain(k) is close to edge i->Chain(j)
                        If Distance(Chain(j), Chain(k)) < JoinDistance Then
                            'Chain(k) is close to Chain(j), they should be combined
                            m = GetEdgeBetween(i, Chain(j)) 'edge i-Chain(j)
                            n = GetEdgeBetween(i, Chain(k)) 'edge i-Chain(k)
                            
                            If m >= 0 And n >= 0 Then
lCheckEdge:
                                'remove any edges from Chain(j) to Chain(k)
                                p = GetEdgeBetween(Chain(j), Chain(k))
                                If p >= 0 Then
                                    Call DelEdge(p)
                                    GoTo lCheckEdge
                                End If
                                
                                Merged = Merged + 1 'at least one change made
                                Nodes(i).mark = 1 'mark node to check again
                                
                                q = CompareRoadtype(Edges(m).roadtype, Edges(n).roadtype)
                                If q = -1 Then
                                    'edge n have higher priority
                                    Call CombineEdges(n, m, i) 'combine edge m into n
                                    Call MergeNodes(Chain(k), Chain(j), 1) 'combine node Chain(j) into Chain(k) w/o moving Chain(k)
                                    Nodes(Chain(k)).mark = 1 'mark node to check once again
                                    Chain(j) = -1 'remove Chain(j) from chain
                                    GoTo lSkipJ 'proceed to next j
                                Else
                                    'edge m have higher priority or edges are equal
                                    Call CombineEdges(m, n, i) 'combine edge n into m
                                    If q = 0 Then
                                        'edges are equal
                                        Call MergeNodes(Chain(j), Chain(k), 0) 'combine with averaging coordinates
                                    Else
                                        'edge m have higher priority
                                        Call MergeNodes(Chain(j), Chain(k), 1) 'combine w/o moving Chain(j)
                                    End If
                                    Nodes(Chain(j)).mark = 1 'mark node to check once again
                                    Chain(k) = -1 'remove Chain(k) from chain
                                    GoTo lSkipK 'proceed to next k
                                End If
                            End If
                        ElseIf Distance(i, Chain(k)) > Dist * AcuteKoeff Then
                            'distance from i to chain(k) is higher than distance from Chain(k) to interval i-Chain(j) in AcuteKoeff times
                            '=> angle Chain(k)-i-Chain(j) < limit angle
                            '=>
                            'Chain(k) should be inserted into edge i-Chain(j)
                            'edge i-Chain(j) became Chain(k)-Chain(j) and keeps all params
                            'edge i-Chain(k) became joined by params
                            m = GetEdgeBetween(i, Chain(j)) 'edge i-Chain(j) - long edge
                            n = GetEdgeBetween(i, Chain(k)) 'edge i-Chain(k) - short edge
                            
                            If m >= 0 And n >= 0 Then
                                If CompareRoadtype(Edges(m).roadtype, Edges(n).roadtype) = -1 Then
                                    'edge n have higher priority
                                Else
                                    'edge m have higher priority or equal
                                    Call ProjectNode(i, Chain(j), Nodes(Chain(k))) 'move Chain(k) to line i-Chain(j)
                                End If
                                Call CombineEdgeParams(n, m, i) 'combine params from m into n
                                Call ReconnectEdge(m, i, Chain(k)) 'edge m is now Chain(k)-Chain(j)
                                Nodes(Chain(j)).mark = 1 'mark nodes as needed to check once again
                                Nodes(Chain(k)).mark = 1
                                Nodes(i).mark = 1
                                Merged = Merged + 1 'at least one change made
                                Chain(j) = -1 'remove Chain(j) from chain, as it is not connected to node i
                                GoTo lSkipJ 'proceed to next j
                            End If
                        End If
                    End If
lSkipK:
                Next
lSkipJ:
            Next
        End If
        
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "JoinAcute (" + CStr(JoinDistance) + ", " + CStr(AcuteKoeff) + ") #" + CStr(PassNumber) + " : " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
            DoEvents
        End If
    Next
    
    If Merged > 0 Then
        'at least one change made - relaunch algorithm
        PassNumber = PassNumber + 1
        Form1.Caption = "JoinAcute (" + CStr(JoinDistance) + ", " + CStr(AcuteKoeff) + ") #" + CStr(PassNumber) + " : merged " + CStr(Merged) 'show progress
        GoTo lIteration
    End If
        
End Sub

'Reconnect edge1 from node1 to node2
'assumed that node1 is present in edge1
Public Sub ReconnectEdge(edge1 As Long, node1 As Long, node2 As Long)
    Dim i As Long
    If Edges(edge1).node1 = node1 Then
        Edges(edge1).node1 = node2
    Else
        Edges(edge1).node2 = node2
    End If
    
    'remove edge1 from node1 edges
    For i = 0 To Nodes(node1).Edges - 1
        If Nodes(node1).edge(i) = edge1 Then
            Nodes(node1).edge(i) = Nodes(node1).edge(Nodes(node1).Edges - 1)
            Nodes(node1).Edges = Nodes(node1).Edges - 1
            GoTo lFound
        End If
    Next
    
lFound:
    'add edge1 to node2 edges
    Call AddEdgeToNode(node2, edge1)
    'Nodes(node2).edge(Nodes(node2).Edges) = edge1
    'Nodes(node2).Edges = Nodes(node2).Edges + 1
End Sub

'Get bounding box of edge
Public Function GetEdgeBbox(edge1 As Long) As bbox
    GetEdgeBbox.lat_min = Nodes(Edges(edge1).node1).lat
    GetEdgeBbox.lat_max = Nodes(Edges(edge1).node1).lat
    GetEdgeBbox.lon_min = Nodes(Edges(edge1).node1).lon
    GetEdgeBbox.lon_max = Nodes(Edges(edge1).node1).lon
    If GetEdgeBbox.lat_min > Nodes(Edges(edge1).node2).lat Then
        GetEdgeBbox.lat_min = Nodes(Edges(edge1).node2).lat
    End If
    If GetEdgeBbox.lat_max < Nodes(Edges(edge1).node2).lat Then
        GetEdgeBbox.lat_max = Nodes(Edges(edge1).node2).lat
    End If
    If GetEdgeBbox.lon_min > Nodes(Edges(edge1).node2).lon Then
        GetEdgeBbox.lon_min = Nodes(Edges(edge1).node2).lon
    End If
    If GetEdgeBbox.lon_max < Nodes(Edges(edge1).node2).lon Then
        GetEdgeBbox.lon_max = Nodes(Edges(edge1).node2).lon
    End If
End Function

'Expand bounding box by distance in metres
Public Function ExpandBbox(ByRef bbox1 As bbox, Dist As Double)
    Dim cos1 As Double
    Dim cos2 As Double
    Dim dist_angle As Double
    dist_angle = RADTODEG * Dist / DATUM_R_OVER 'distance in degrees of latitude
    bbox1.lat_min = bbox1.lat_min - dist_angle
    bbox1.lat_max = bbox1.lat_max + dist_angle
    
    cos1 = Cos(bbox1.lat_min * DEGTORAD)
    cos2 = Cos(bbox1.lat_max * DEGTORAD)
    If cos2 < cos1 Then cos1 = cos2 'smallest cos() - further from equator
    If cos1 < 0.01 Then cos1 = 0.01 'beyond 89' lat, ex. Antarctic Territories
    dist_angle = dist_angle / cos1 'distance in degrees of longtitue
    bbox1.lon_min = bbox1.lon_min - dist_angle
    bbox1.lon_max = bbox1.lon_max + dist_angle
    'TODO: fix (not safe to 180/-180 edge)
    
End Function


'Join two directions of road way
'MaxCosine - cosine of max angle between start edges, -0.996 means (175,180) degrees - contradirectional edges or close
'MaxCosine2 - cosine of max angle between other edges, during going-by-two-ways
'MinChainLen - length of min two-way road to join
Public Sub JoinDirections3(JoinDistance As Double, MaxCosine As Double, MaxCosine2 As Double, MinChainLen As Double, CombineDistance As Double)
    Dim i As Long, j As Long, k As Long, mode1 As Long
    Dim e As Long, d As Long, q As Long
    Dim dist1 As Double, dist2 As Double
    Dim bbox_edge As bbox
    Dim angl As Double
    Dim min_dist As Double, min_dist_edge As Long
    Dim roadtype As Long
    Dim speednew As Integer
    
    Dim EdgesForw() As Long 'chain of forward edges
    Dim EdgesBack() As Long 'chain of backward edges
    Dim LoopChain As Long '1 if road is circled
    Dim HalfChain As Long 'len of half of road
    
    'Algorithm will check all non-link oneway edges for presence of contradirectional edge in the vicinity
    'All found pairs of edges will be checked in both directions by GoByTwoWays function
    'for presence of continuous road of one type
    'During this check will be created new chain of nodes, which is projection of joining nodes into middle line
    'Then both found ways will be joined into one bidirectional way, consist from new nodes
    'All related roads will reconnected to new way and old edges were deleted
    
    'mark all nodes as not checked
    For i = 0 To NodesNum - 1
        Nodes(i).mark = -1 'not moved
    Next
    For i = 0 To EdgesNum - 1
        Edges(i).mark = 1
        If Edges(i).node1 = -1 Or Edges(i).oneway = 0 Then GoTo lFinMarkEdge 'skip deleted and 2-ways edges
        If (Edges(i).roadtype And HIGHWAY_MASK_LINK) > 0 Then GoTo lFinMarkEdge 'skip links
        If Nodes(Edges(i).node1).Edges <> 2 And Nodes(Edges(i).node2).Edges <> 2 Then GoTo lFinMarkEdge 'skip edges between complex connections
        Edges(i).mark = 0
lFinMarkEdge:
    Next
    
    'rebuild cluster-index from 0
    Call BuildNodeClusterIndex(0)
    
    For i = 0 To EdgesNum - 1
        If Edges(i).mark > 0 Or Edges(i).node1 < 0 Then GoTo lSkipEdge 'skip marked edge or deleted
        bbox_edge = GetEdgeBbox(i) 'get bbox
        Call ExpandBbox(bbox_edge, JoinDistance) 'expand it
        min_dist = JoinDistance
        min_dist_edge = -1
        
        mode1 = 0 'first
lSkipNode2:
        k = GetNodeInBboxByCluster(bbox_edge, mode1)
        mode1 = 1 'next (next time)
        If k = -1 Then GoTo lAllNodes 'no more nodes
        
        If k = Edges(i).node1 Or k = Edges(i).node2 Or Nodes(k).NodeID = MARK_NODEID_DELETED Or Nodes(k).Edges <> 2 Then GoTo lSkipNode2 'skip nodes of same edge, deleted and complex nodes
        
        dist1 = DistanceToSegment(Edges(i).node1, Edges(i).node2, k) 'calc dist from found node to our edge
        If dist1 > min_dist Then GoTo lSkipNode2 'too far, skip
        
        'node is on join distance, check all (2) edges
        For d = 0 To 1
            q = Nodes(k).edge(d)
            If Edges(q).node1 = -1 Or Edges(q).oneway = 0 Or Edges(q).roadtype <> Edges(i).roadtype Then GoTo lSkipEdge2 'deleted or 2-way edge or other road class
            angl = CosAngleBetweenEdges(q, i)
            If angl < MaxCosine Then
                'contradirectional edge or close
                
                dist1 = DistanceBetweenSegments(i, q)
                If dist1 < min_dist Then min_dist = dist1: min_dist_edge = q 'found edge close enough
            End If
lSkipEdge2:
        Next

        GoTo lSkipNode2:
        
lAllNodes:
        'all nodes in bbox check
        
        If min_dist_edge > -1 Then
            'found edge close enough
            'now - trace two ways in both directions
            'in the process we will fill Chain array with all nodes of joining ways
            'sequence of nodes in Chain will correspond to sequence of nodes on combined way
            'index of new node, where old node should join, will in .mark field of old nodes
            'also will be created two lists of deleting edges separated to two directions - arrays TWforw and TWback
            
            roadtype = Edges(i).roadtype
            LoopChain = 0
            ChainNum = 0
            TWforwNum = 0
            TWbackNum = 0
            
            'first pass, in direction of edge i
            Call GoByTwoWays(i, min_dist_edge, JoinDistance, CombineDistance, MaxCosine2, 0)
            
            'reverse of TWforw and TWback removed, as order of edges have no major effect
            
            'reverse Chain
            Call ReverseArray(Chain, ChainNum)
            
            'second pass, in direction of min_dist_edge
            Call GoByTwoWays(min_dist_edge, i, JoinDistance, CombineDistance, MaxCosine2, 1)
            
            If Chain(0) = Chain(ChainNum - 1) Then LoopChain = 1 'first and last nodes coincide - this is loop road
            
            HalfChain = ChainNum / 2 'half len of road in nodes
            If HalfChain < 10 Then HalfChain = ChainNum + 1 'will "kill" halfchain limit for very short loops
            
            'call metric length of found road
            dist1 = 0
            For j = 1 To ChainNum - 1
                dist1 = dist1 + Distance(Nodes(Chain(j - 1)).mark, Nodes(Chain(j)).mark)
            Next
            
            If dist1 < MinChainLen Then
                'road is too short -> unmark all edges and not delete anything
                For j = 0 To ChainNum - 1
                    For k = 0 To Nodes(Chain(j)).Edges - 1
                        If Edges(Nodes(Chain(j)).edge(k)).mark = 2 Then
                            Edges(Nodes(Chain(j)).edge(k)).mark = 1
                        End If
                    Next
                Next
                GoTo lSkipDel
            End If
            
            
            'process both directions edges list
            'to build index of edges, which joins between each pair of nodes in Chain
            
            'note: is some rare cases chain of nodes have pleats, where nodes of one directions
            'in Chain swaps position due to non uniform projecting of nodes to middle-line
            'In this cases one or more edges joins to bidirectional road in backward direction
            'to the original direction of this one-way line
            'These edges could be ignored during combining parameter of bidirectional road
            '(as they are usually very short)
            'also at least two other edges will overlap in at least one interval between nodes
            'only one of them will be counted during combining parameter (last in TW* array)
            'this is considired acceptable, as they are near edges of very same road
            
            ReDim EdgesForw(ChainNum)
            ReDim EdgesBack(ChainNum)
            For j = 0 To ChainNum - 1
                EdgesForw(j) = -1
                EdgesBack(j) = -1
            Next
            
            'process forward direction
            For j = 0 To TWforwNum - 1
                e = Edges(TWforw(j)).node1
                d = Edges(TWforw(j)).node2
                e = FindInChain(e) 'get indexes of nodes inside Chain
                d = FindInChain(d)
                If e = -1 Or d = -1 Then
                    '(should not happen)
                    'edge with nodes not in chain - skip
                    GoTo lSkip1
                End If
                
                If e < d Then
                    'normal forward edge (or pleat crossing 0 of chain)
                    ' ... e ---> d .....
                    If LoopChain = 1 And (d - e) > HalfChain Then GoTo lSkip1: 'skip too long edges on loop chains as it could be wrong (i.e. pleat edge which cross 0 of chain)
                    For q = e To d - 1
                        'in forward direction between q and q+1 node is edge TWforw(j)
                        EdgesForw(q) = TWforw(j)
                    Next
                Else
                    'pleat edge (or normal crossing 0 of chain)
                    ' ---.---> d .... ... .... e --->
                    If LoopChain = 0 Then GoTo lSkip1 'on straight chains forward edge could not go backward without pleat
                    If (e - d) > HalfChain Then
                        'e and d is close to ends of chain
                        '-> this is really forward edge crossing 0 of chain in a loop road
                        For q = 0 To d - 1
                            EdgesForw(q) = TWforw(j)
                        Next
                        For q = e To ChainNum - 1
                            EdgesForw(q) = TWforw(j)
                        Next
                    End If
                End If
lSkip1:
            Next
            
            'process backward direction
            For j = 0 To TWbackNum - 1
                e = Edges(TWback(j)).node1
                d = Edges(TWback(j)).node2
                e = FindInChain(e) 'get indexes of nodes inside Chain
                d = FindInChain(d)
                If e = -1 Or d = -1 Then
                    '(should not happen)
                    'edge with nodes not in chain - skip
                    GoTo lSkip2
                End If
                
                If d < e Then
                    'normal backward edge (or pleat crossing 0 of chain)
                    ' ... d <--- e .....
                    If LoopChain = 1 And (e - d) > HalfChain Then GoTo lSkip2: 'skip too long edges on loop chains as it could be wrong (i.e. pleat edge which cross 0 of chain)
                    For q = d To e - 1
                        EdgesBack(q) = TWback(j)
                    Next
                Else
                    'pleat edge (or normal crossing 0 of chain)
                    ' <-.-- e ... ... .... ... d <--.---.---
                    If LoopChain = 0 Then GoTo lSkip2 'on straight chains backward edge could not go forward without pleat
                    If (d - e) > HalfChain Then
                        'e and d is close to ends of chain
                        '-> this is really backward edge crossing 0 of chain in a loop road
                        For q = 0 To e - 1
                            EdgesBack(q) = TWback(j)
                        Next
                        For q = d To ChainNum - 1
                            EdgesBack(q) = TWback(j)
                        Next
                    End If
                End If
lSkip2:
            Next
            
            For j = 1 To ChainNum - 1
                d = Nodes(Chain(j - 1)).mark
                e = Nodes(Chain(j)).mark
                If d <> e Then
                    k = JoinByEdge(Nodes(Chain(j - 1)).mark, Nodes(Chain(j)).mark)
                    Edges(k).roadtype = roadtype
                    Edges(k).oneway = 0
                    Edges(k).mark = 1
                    If EdgesForw(j - 1) = -1 And EdgesBack(j - 1) = -1 Then
                        'no edges for this interval between nodes
                        '(should never happens)
                        Edges(k).speed = 3 'default value
                        Edges(k).label = ""
                    Else
                        'get minimal speed class of both edges
                        speednew = 10
                        Call ResetLabelStats
                        If EdgesForw(j - 1) <> -1 Then
                            'forward edge present
                            speednew = Edges(EdgesForw(j - 1)).speed
                            Call AddLabelStat0(Edges(EdgesForw(j - 1)).label)
                        End If
                        If EdgesBack(j - 1) <> -1 Then
                            'backward edge present
                            If speednew > Edges(EdgesBack(j - 1)).speed Then speednew = Edges(EdgesBack(j - 1)).speed
                            Call AddLabelStat0(Edges(EdgesBack(j - 1)).label)
                        End If
                        Edges(k).speed = speednew
                        Edges(k).label = GetLabelByStats(0)
                    End If
                    
                    
                    'ends of chain could be oneway if only one edge (or even part is joining there
                    'ex:     * ------> * --------> * ----------> *
                    '             * <-------- * <--------- * <---------- *
                    'joins into:
                    '        *--->*----*------*----*-------*-----*<------*
                    
                    If EdgesBack(j - 1) = -1 Then
                        'no backward edge - result in one-way
                        Edges(k).oneway = 1
                    ElseIf EdgesForw(j - 1) = -1 Then
                        'no forward edge - result in one-way, backward to other road
                        Edges(k).oneway = 1
                        Edges(k).node1 = Nodes(Chain(j)).mark
                        Edges(k).node2 = Nodes(Chain(j - 1)).mark
                    End If
                End If
            Next
            
            'delete all old edges
            For j = 0 To TWforwNum - 1
                Call DelEdge(TWforw(j))
            Next
            For j = 0 To TWbackNum - 1
                Call DelEdge(TWback(j))
            Next
            
            'merge all old nodes into new ones
            For j = 0 To ChainNum - 1
                Call MergeNodes(Nodes(Chain(j)).mark, Chain(j), 1)
            Next
            
            'update cluster index to include only newly created nodes (i.e. nodes of joined road)
            Call BuildNodeClusterIndex(1)
            
lSkipDel:
        End If
        
        Edges(i).mark = 1 'mark edge as checked
        
lSkipEdge:

        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "JoinDirections3 (" + CStr(JoinDistance) + ", " + CStr(MaxCosine) + ", " + CStr(MaxCosine2) + ", " + CStr(MinChainLen) + ", " + CStr(CombineDistance) + ") : " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
            DoEvents
        End If
    Next
    

End Sub


'Reverse array into backward direction
Public Sub ReverseArray(ByRef Arr() As Long, Num As Long)
    Dim i As Long
    Dim j As Long
    Dim t As Long
    j = Num \ 2 'half of len
    For i = 0 To j - 1
        'swap elements from first and second halfs
        t = Arr(i)
        Arr(i) = Arr(Num - 1 - i)
        Arr(Num - 1 - i) = t
    Next
End Sub

'Find edges of two way road
'Algorithm goes by finding next edge on side, which is not leading
'Found new node (end of found edge) is projected to local middle line
'Array Chain is filled by found nodes
'Arrays TWforw and TWback is filled by found edges
'
'edge1,edge2 - start edges
'JoinDistance - distance limit between two ways
'CombineDistance - distance to join two nodes into one (on middle line)
'MaxCosine2 - angle limit between edges
'Params: 0 - first pass (chain empty, go by edge1 direction)
'        1 - second pass (chain contains all 4 nodes of edges at the end, go by edge2 direction)
Public Sub GoByTwoWays(edge1 As Long, edge2 As Long, JoinDistance As Double, CombineDistance As Double, MaxCosine2 As Double, Params As Long)
    Dim i As Long, j As Long, k As Long
    
    Dim edge_side1 As Long 'arrow-head edges
    Dim edge_side2 As Long
    Dim side1i As Long, side1j As Long 'arrow-head nodes
    Dim side2i As Long, side2j As Long
    Dim side1circled As Long 'flags of circle on each side
    Dim side2circled As Long
    
    Dim side(4) As Long
    Dim Dist(4) As Double
    Dim dist_t As Double
    Dim dx As Double
    Dim dy As Double
    Dim px As Double, py As Double
    Dim dd As Double
    Dim roadtype As Long
    Dim angl As Double
    Dim calc_side As Long
    Dim angl_min As Double, angl_min_edge As Long
    Dim checkchain As Long
    Dim PassNumber As Long

    roadtype = Edges(edge1).roadtype 'keep road type for comparing
    
    Edges(edge1).mark = 2 'mark edges as participating in joining
    Edges(edge2).mark = 2
    
    edge_side1 = edge1 'arrow-head of finding chains
    edge_side2 = edge2
    
    'i node is back, j is front of arrow-head - on both sides
    side1i = Edges(edge1).node1
    side1j = Edges(edge1).node2
    side2i = Edges(edge2).node2
    side2j = Edges(edge2).node1
    
    side1circled = 0 'circles not yet found
    side2circled = 0
    
    PassNumber = 0
    If (Params And 1) Then PassNumber = 1
    
    If PassNumber = 1 Then
        'second pass
        'skip initial part, as it is already done in first pass
        GoTo lKeepGoing
    End If
    
    'middle line projection vector
    'TODO: fix (not safe to 180/-180 edge)
    dx = (Nodes(side1j).lat - Nodes(side1i).lat) + (Nodes(side2j).lat - Nodes(side2i).lat) 'sum of two edges
    dy = (Nodes(side1j).lon - Nodes(side1i).lon) + (Nodes(side2j).lon - Nodes(side2i).lon)
    px = (Nodes(side1i).lat + Nodes(side2i).lat) * 0.5 'start point - average of two starts
    py = (Nodes(side1i).lon + Nodes(side2i).lon) * 0.5

    side(0) = side1i
    side(1) = side1j
    side(2) = side2i
    side(3) = side2j

    'calc relative positions of projections of all 4 noes to edge1
    dd = 1 / (dx * dx + dy * dy)
    For i = 0 To 3
        Dist(i) = (Nodes(side(i)).lat - px) * dx + (Nodes(side(i)).lon - py) * dy
    Next

    'Sort dist() and side() by dist() by bubble sort
    For i = 0 To 3
        For j = i + 1 To 3
            If Dist(j) < Dist(i) Then
                dist_t = Dist(j): Dist(j) = Dist(i): Dist(i) = dist_t
                k = side(j): side(j) = side(i): side(i) = k
            End If
        Next
    Next
    
    'Add nodes to chain in sorted order
    For i = 0 To 3
        Call AddChain(side(i))
        Nodes(NodesNum).Edges = 0
        Nodes(NodesNum).NodeID = -1
        Nodes(NodesNum).mark = -1
        Nodes(side(i)).mark = NodesNum 'info that old node will collapse to this new one
        Nodes(NodesNum).lat = px + Dist(i) * dx * dd 'projected coordinates
        Nodes(NodesNum).lon = py + Dist(i) * dy * dd
        Call AddNode
    Next
    
lKeepGoing:

    angl_min = MaxCosine2: angl_min_edge = -1
    
    If Chain(ChainNum - 1) = side1j Then
        'side1 is leading, side2 should be prolonged
        calc_side = 2
    Else
        'side2 is leading, side1 should be prolonged
        calc_side = 1
    End If
        
    If calc_side = 2 Then
        'search edge from side2j which is most opposite to edge_side1
        For i = 0 To Nodes(side2j).Edges - 1
            j = Nodes(side2j).edge(i)
            If j = edge_side2 Or Edges(j).node1 < 0 Or Edges(j).oneway = 0 Or Edges(j).roadtype <> roadtype Or Edges(j).node2 <> side2j Then GoTo lSkipEdgeSide2
                'skip same edge_side2, deleted, 2-ways, other road types and directed from this node outside
                dist_t = DistanceBetweenSegments(j, edge_side1)
                If dist_t > JoinDistance Then GoTo lSkipEdgeSide2 'skip too far edges
                angl = CosAngleBetweenEdges(j, edge_side1)
                If angl < angl_min Then angl_min = angl: angl_min_edge = j 'remember edge with min angle
lSkipEdgeSide2:
        Next
        
        Edges(edge_side2).mark = 2 'mark edge as participating in joining
        Call AddTW(edge_side2, PassNumber) 'add edge to chain (depending on pass number)
        
        If angl_min_edge = -1 Then
            'no edge found - end of chain
            Edges(edge_side1).mark = 2 'mark last edge of side1
            Call AddTW(edge_side1, 1 - PassNumber) 'and add it to chain
            GoTo lChainEnds
        End If
        
        edge_side2 = angl_min_edge
        side2i = side2j 'update i and j nodes of side
        side2j = Edges(edge_side2).node1
        
        If Edges(edge_side2).mark = 2 Then
            'found marked edge, this means that we found cycle
            side2circled = 1
        End If
        
        If side2j = side1j Then
            'found joining of two directions, should end chain
            Edges(edge_side2).mark = 2 'mark both last edges as participating in joining
            Edges(edge_side1).mark = 2
            Call AddTW(edge_side2, PassNumber) 'add them to chains
            Call AddTW(edge_side1, 1 - PassNumber)
            GoTo lChainEnds
        End If
    
    Else
        'search edge from side1j which is most opposite to edge_side2
        For i = 0 To Nodes(side1j).Edges - 1
            j = Nodes(side1j).edge(i)
            If j = edge_side1 Or Edges(j).oneway = 0 Or Edges(j).roadtype <> roadtype Or Edges(j).node1 <> side1j Then GoTo lSkipEdgeSide1
                'skip same edge_side1, 2-ways, other road types and directed from this node outside
                dist_t = DistanceBetweenSegments(j, edge_side2)
                If dist_t > JoinDistance Then GoTo lSkipEdgeSide1 'skip too far edges
                angl = CosAngleBetweenEdges(j, edge_side2)
                If angl < angl_min Then angl_min = angl: angl_min_edge = j 'remember edge with min angle
lSkipEdgeSide1:
        Next
        
        Edges(edge_side1).mark = 2 'mark edge as participating in joining
        Call AddTW(edge_side1, 1 - PassNumber) 'add edge to chain (depending on pass number)
        
        If angl_min_edge = -1 Then
            'no edge found - end of chain
            Edges(edge_side2).mark = 2 'mark last edge of side2
            Call AddTW(edge_side2, PassNumber) 'and add it to chain
            GoTo lChainEnds
        End If
        
        edge_side1 = angl_min_edge
        side1i = side1j 'update i and j nodes of side
        side1j = Edges(edge_side1).node2
    
        If Edges(edge_side1).mark = 2 Then
            'found marked edge, means, that we found cycle
            side1circled = 1
        End If
        
        If side2j = side1j Then
            'found marked edge, this means that we found cycle
            Edges(edge_side2).mark = 2 'mark both last edges as participating in joining
            Edges(edge_side1).mark = 2
            Call AddTW(edge_side2, PassNumber) 'add them to chains
            Call AddTW(edge_side1, 1 - PassNumber)
            GoTo lChainEnds
        End If
    End If

    'middle line projection vector
    'TODO: fix (not safe to 180/-180 edge)
    dx = Nodes(side1j).lat - Nodes(side1i).lat + Nodes(side2j).lat - Nodes(side2i).lat
    dy = Nodes(side1j).lon - Nodes(side1i).lon + Nodes(side2j).lon - Nodes(side2i).lon
    px = (Nodes(side1i).lat + Nodes(side2i).lat) * 0.5
    py = (Nodes(side1i).lon + Nodes(side2i).lon) * 0.5
    dd = 1 / (dx * dx + dy * dy)
    
    checkchain = ChainNum 'remember current chain len
    
    If calc_side = 2 Then
        'project j node from side2 to middle line
        dist_t = (Nodes(side2j).lat - px) * dx + (Nodes(side2j).lon - py) * dy
        Call AddChain(side2j)
        Nodes(side2j).mark = NodesNum 'old node will collapse to this new one
    Else
        'project j node from side1 to middle line
        dist_t = (Nodes(side1j).lat - px) * dx + (Nodes(side1j).lon - py) * dy
        Call AddChain(side1j)
        Nodes(side1j).mark = NodesNum 'old node will collapse to this new one
    End If
   
    'create new node
    Nodes(NodesNum).Edges = 0
    Nodes(NodesNum).NodeID = -1
    Nodes(NodesNum).mark = -1
    Nodes(NodesNum).lat = px + dist_t * dx * dd
    Nodes(NodesNum).lon = py + dist_t * dy * dd
    
    'reproject prev node into current middle line ("ChainNum - 2" because ChainNum were updated above by AddChain)
    j = Nodes(Chain(ChainNum - 2)).mark
    dist_t = (Nodes(j).lat - px) * dx + (Nodes(j).lon - py) * dy
    Nodes(j).lat = px + dist_t * dx * dd
    Nodes(j).lon = py + dist_t * dy * dd
    
    If Distance(j, NodesNum) < CombineDistance Then
        'Distance from new node to prev-one is too small, collapse node with prev-one
        'TODO(?): averaging coordinates?
        If calc_side = 2 Then
            Nodes(side2j).mark = j
        Else
            Nodes(side1j).mark = j
        End If
        'do not call AddNode -> new node will die
    Else
        Call AddNode
        Call FixChainOrder(checkchain) 'fix order of nodes in chain
    End If

    If side1circled > 0 And side2circled > 0 Then GoTo lFoundCycle 'both sides circled - whole road is a loop
    
    GoTo lKeepGoing 'proceed to searching next edge
    
lChainEnds:

    'Node there is chance, that circular way will be not closed from one of sides
    'Algorithm does not handle this case, it should collapse during juctions collapsing

    Exit Sub
    
lFoundCycle:
    'handle cycle road
    
    'find all nodes from end of chain which is present in chain two times
    'remove all of them, except last one
    'in good cases last node should be same as first node
    'TODO: what if not?
    For i = ChainNum - 1 To 0 Step -1
        For j = 0 To i - 1
            If Chain(i) = Chain(j) Then GoTo lFound
        Next
        'not found
        ChainNum = i + 2 'keep this node (which is one time in chain) and next one (which is two times)
        GoTo lExit
lFound:
    Next
    
lExit:
    
End Sub


'Fix order of nodes in Chain
'Fixing is needed when last node is not new arrow-head of GoByTwoWays algorithm (ex. several short edges of one side, but long edge of other side)
Public Sub FixChainOrder(checkindex As Long)
    
    Dim i2 As Long, i1 As Long, i0 As Long, k As Long
    Dim p As Double
    If checkindex < 2 Then Exit Sub '2 or less nodes in chain, nothing to fix
    
    i2 = Nodes(Chain(checkindex)).mark 'last new node
    If i2 < 0 Then Exit Sub 'exit in case of probles
    i1 = Nodes(Chain(checkindex - 1)).mark 'prev new node
    If i1 < 0 Then Exit Sub
    i0 = Nodes(Chain(checkindex - 2)).mark 'prev-prev new node
    If i0 < 0 Then Exit Sub
    
    k = 3
    'if prev-prev new nodes is combined with prev new node - find diffent node backward
    While i0 = i1
        If checkindex < k Then Exit Sub 'reach Chain(0)
        i0 = Nodes(Chain(checkindex - k)).mark
        k = k + 1
    Wend
    
    'Scalar multiplication of vectors i0->i1 and i1->i2
    p = (Nodes(i2).lat - Nodes(i1).lat) * (Nodes(i1).lat - Nodes(i0).lat) + _
        (Nodes(i2).lon - Nodes(i1).lon) * (Nodes(i1).lon - Nodes(i0).lon)
        
    If p < 0 Then
        'vectors are contradirectional -> swap
        i0 = Chain(checkindex)
        Chain(checkindex) = Chain(checkindex - 1)
        Chain(checkindex - 1) = i0
        'check last new node on new place
        Call FixChainOrder(checkindex - 1)
    End If
End Sub

'Build cluster index
'Cluster index allow to quickly find nodes in specific bbox
'Cluster index is collections of nodes chains, where starts can be selected from coordinates
'and continuation - by indexes in chains
'Flags: 1 - only update from ClustersIndexedNodes to NodesNum (0 - full re/build)
Public Sub BuildNodeClusterIndex(Flags As Long)
    Dim i As Long, j As Long, k As Long
    Dim x As Long
    Dim y As Long
    
    If (Flags And 1) <> 0 Then
        'Only update
        'TODO(?): remove chain from deleted nodes
        ReDim Preserve ClustersChain(NodesNum)
        GoTo lClustering
    End If
    
    'calc overall bbox
    Dim wholeBbox As bbox
    wholeBbox.lat_max = -360
    wholeBbox.lat_min = 360
    wholeBbox.lon_max = -360
    wholeBbox.lon_min = 360
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED Then 'skip deleted nodes
            If Nodes(i).lat < wholeBbox.lat_min Then wholeBbox.lat_min = Nodes(i).lat
            If Nodes(i).lat > wholeBbox.lat_max Then wholeBbox.lat_max = Nodes(i).lat
            If Nodes(i).lon < wholeBbox.lon_min Then wholeBbox.lon_min = Nodes(i).lon
            If Nodes(i).lon > wholeBbox.lon_max Then wholeBbox.lon_max = Nodes(i).lon
        End If
    Next
    
    ClustersIndexedNodes = 0
    If wholeBbox.lat_max < wholeBbox.lat_min Or wholeBbox.lon_max < wholeBbox.lon_min Then Exit Sub 'no nodes at all or something wrong
    
    'calc number of clusters
    ClustersLatNum = 1 + (wholeBbox.lat_max - wholeBbox.lat_min) / Control_ClusterSize
    ClustersLonNum = 1 + (wholeBbox.lon_max - wholeBbox.lon_min) / Control_ClusterSize
    
    ReDim ClustersFirst(ClustersLatNum * ClustersLonNum) 'starts of chains
    ReDim ClustersLast(ClustersLatNum * ClustersLonNum)  'ends of chains (for updating)
    ReDim ClustersChain(NodesNum) 'whole chain
    
    ClustersLat0 = wholeBbox.lat_min 'edge of overall bbox
    ClustersLon0 = wholeBbox.lon_min
    
    For i = 0 To ClustersLatNum * ClustersLonNum - 1
        ClustersFirst(i) = -1 'no nodes in cluster yet
        ClustersLast(i) = -1
    Next
    
    ClustersIndexedNodes = 0
    
lClustering:
    For i = ClustersIndexedNodes To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED Then
            'get cluster from lat/lon
            x = (Nodes(i).lat - ClustersLat0) / Control_ClusterSize
            y = (Nodes(i).lon - ClustersLon0) / Control_ClusterSize
            j = x + y * ClustersLatNum
            
            k = ClustersLast(j)
            If k = -1 Then
                'first index in chain of this cluster
                ClustersFirst(j) = i
            Else
                'continuing chain
                ClustersChain(k) = i
            End If
            ClustersChain(i) = -1 'this is last node in chain
            ClustersLast(j) = i
        End If
    Next

    ClustersIndexedNodes = NodesNum 'last node in cluster index

End Sub


'Find node in bbox by using cluster index
'Flags: 1 - next (0 - first)
Public Function GetNodeInBboxByCluster(box1 As bbox, Flags As Long) As Long
    Dim i As Long, j As Long, k As Long
    Dim x As Long, y As Long
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    
    If (Flags And 1) = 0 Then
        'first node needed
        
        'get coordinates of all needed clusters
        x1 = (box1.lat_min - ClustersLat0) / Control_ClusterSize
        x2 = (box1.lat_max - ClustersLat0) / Control_ClusterSize
        y1 = (box1.lon_min - ClustersLon0) / Control_ClusterSize
        y2 = (box1.lon_max - ClustersLon0) / Control_ClusterSize
        
        If x1 < 0 Then x1 = 0
        If x2 < 0 Then x2 = 0
        If x1 >= ClustersLatNum Then x1 = ClustersLatNum - 1
        If x2 >= ClustersLatNum Then x2 = ClustersLatNum - 1
        If y1 < 0 Then y1 = 0
        If y2 < 0 Then y2 = 0
        If y1 >= ClustersLonNum Then y1 = ClustersLonNum - 1
        If y2 >= ClustersLonNum Then y2 = ClustersLonNum - 1
        
        ClustersFindLastBbox = box1 'store bbox for next searches
        x = x1
        y = y1
        GoTo lCheckFirst
    End If
    
    If ClustersFindLastNode = -1 Then
        'Last time nothing found - nothing to do further
        GetNodeInBboxByCluster = -1
        GoTo lExit
    End If
    
    'get coordinates of all needed clusters
    x1 = (ClustersFindLastBbox.lat_min - ClustersLat0) / Control_ClusterSize
    x2 = (ClustersFindLastBbox.lat_max - ClustersLat0) / Control_ClusterSize
    y1 = (ClustersFindLastBbox.lon_min - ClustersLon0) / Control_ClusterSize
    y2 = (ClustersFindLastBbox.lon_max - ClustersLon0) / Control_ClusterSize
    
    If x1 < 0 Then x1 = 0
    If x2 < 0 Then x2 = 0
    If x1 >= ClustersLatNum Then x1 = ClustersLatNum - 1
    If x2 >= ClustersLatNum Then x2 = ClustersLatNum - 1
    If y1 < 0 Then y1 = 0
    If y2 < 0 Then y2 = 0
    If y1 >= ClustersLonNum Then y1 = ClustersLonNum - 1
    If y2 >= ClustersLonNum Then y2 = ClustersLonNum - 1

    'get coordinates of last used cluster
    x = ClustersFindLastCluster Mod ClustersLatNum
    y = (ClustersFindLastCluster - x) \ ClustersLatNum
    
lNextNode:
    k = ClustersChain(ClustersFindLastNode) 'get node from chain
    If k <> -1 Then
        'not end of chain
lCheckNode:
        ClustersFindLastNode = k 'keep as last result
        If Nodes(k).NodeID = MARK_NODEID_DELETED Then GoTo lNextNode 'deleted node - find next
        If Nodes(k).lat < ClustersFindLastBbox.lat_min Or Nodes(k).lat > ClustersFindLastBbox.lat_max Then GoTo lNextNode 'node outside desired bbox - find next
        If Nodes(k).lon < ClustersFindLastBbox.lon_min Or Nodes(k).lon > ClustersFindLastBbox.lon_max Then GoTo lNextNode
        GetNodeInBboxByCluster = k 'OK, found
        GoTo lExit
    End If
    
    'end of chain -> last node in cluster
    
lNextCluster:
    'proceed to next cluster
    
    x = x + 1
    If x > x2 Then y = y + 1: x = x1 'next line of cluster
    If y > y2 Then
        'last cluster - no nodes
        GetNodeInBboxByCluster = -1 'nothing found
        ClustersFindLastNode = -1 'nothing will be found next time
        ClustersFindLastCluster = -1
        GoTo lExit
    End If
    
lCheckFirst:
    'get first node of cluster
    
    j = x + y * ClustersLatNum
    k = ClustersFirst(j)
    If k = -1 Then GoTo lNextCluster 'no first node - skip cluster
    ClustersFindLastCluster = j
    GoTo lCheckNode 'there is first node - check it
    
lExit:
    
End Function


'Remove all labels stats from memory
Public Sub ResetLabelStats()
    LabelStatsNum = 0
End Sub


'Add label to label stats
Public Sub AddLabelStat0(Text As String)
    'Call AddLabelStat1(Text) 'majoritary version
    Call AddLabelStat2(Text) 'combinatory version
End Sub

Public Function GetLabelByStats(Flags As Long)
    'GetLabelByStats = GetLabelByStats1(Text) 'majoritary version
    GetLabelByStats = GetLabelByStats2(0) 'combinatory version
End Function


'Add label completely (for majoritary calc and so on)
Public Sub AddLabelStat1(Text As String)
    Dim i As Long
    If Len(Text) < 1 Then Exit Sub 'skip empty strings
    For i = 0 To LabelStatsNum - 1
        If LabelStats(i).Text = Text Then
            'already present - increment count
            LabelStats(i).count = LabelStats(i).count + 1
            Exit Sub
        End If
    Next
    'not present - add
    LabelStats(LabelStatsNum).Text = Text
    LabelStats(LabelStatsNum).count = 1
    LabelStatsNum = LabelStatsNum + 1
    If LabelStatsNum >= LabelStatsAlloc Then
        'realloc if needed
        LabelStatsAlloc = LabelStatsAlloc * 2
        ReDim Preserve LabelStats(LabelStatsAlloc)
    End If
End Sub


'Add label by parts (for combinatory calc and so on)
Public Sub AddLabelStat2(Text As String)
    Dim i As Long
    Dim marks() As String
    If Len(Text) < 1 Then Exit Sub
    marks = Split(Text, ",") 'split string by delimiter into set of strings
    For i = 0 To UBound(marks)
        Call AddLabelStat1(marks(i))
    Next
End Sub


'Get label of road from stats, combinatory version
'Flags: 0 - default
Public Function GetLabelByStats2(Flags As Long) As String
    Dim i As Long
    
    GetLabelByStats2 = ""
    If LabelStatsNum = 0 Then Exit Function 'no labels in stats
    
    'combine all labels from stats
    GetLabelByStats2 = LabelStats(0).Text
    For i = 1 To LabelStatsNum - 1
        GetLabelByStats2 = GetLabelByStats2 + "," + LabelStats(i).Text
    Next
    
End Function

'
'Get label of road from stats, majoritary version
'Flags: 0 - default
Public Function GetLabelByStats1(Flags As Long) As String
    Dim i As Long
    Dim max_count As Long, max_len As Long, max_index As Long

    'select label with max count and max len (amoung max count)
    max_count = -1
    max_len = -1
    max_index = -1
    For i = 0 To LabelStatsNum - 1
        If LabelStats(i).count > max_count Then
            'max count
            max_count = LabelStats(i).count
            max_index = i
            max_len = Len(LabelStats(i).Text)
        ElseIf LabelStats(i).count = max_count Then
            If Len(LabelStats(i).Text) > max_len Then
                'max len amoung max count
                max_count = LabelStats(i).count
                max_index = i
                max_len = Len(LabelStats(i).Text)
            End If
        End If
    Next
    If max_index > -1 Then
        GetLabelByStats1 = LabelStats(max_index).Text
    Else
        'nothing found
        GetLabelByStats1 = ""
    End If

End Function



'Add edge into one of TW arrays
'side: 0 - into TWforw, 1 - into TWback
Public Function AddTW(edge1 As Long, side As Long)
    If side = 1 Then
        TWback(TWbackNum) = edge1
        TWbackNum = TWbackNum + 1
        If TWbackNum >= TWalloc Then GoTo lRealloc
    Else
        TWforw(TWforwNum) = edge1
        TWforwNum = TWforwNum + 1
        If TWforwNum >= TWalloc Then
lRealloc:
            'realloc if needed
            TWalloc = TWalloc * 2
            ReDim Preserve TWforw(TWalloc)
            ReDim Preserve TWback(TWalloc)
        End If
    End If
End Function


'Collapse all edges shorter than CollapseDistance (also will kill void edges)
'Will collapse edges one by one, so should be called somewhere in the end of optimization
Public Sub CollapseShortEdges(CollapseDistance As Double)
    Dim i As Long, j As Long, k As Long
    Dim somedeleted As Long
    Dim EdgeLen As Double

lIteration:
    somedeleted = 0
    For i = 0 To EdgesNum - 1
        If Edges(i).node1 >= 0 Then
            EdgeLen = Distance(Edges(i).node1, Edges(i).node2)
            If EdgeLen < CollapseDistance Then
                j = Edges(i).node1
                k = Edges(i).node2
                Call DelEdge(i) 'del this edge
                If j <> k Then Call MergeNodes(j, k, 0) 'merge nodes, only if they are different
                somedeleted = 1
            End If
        End If
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "CollapseShortEdges (" + CStr(CollapseDistance) + ") " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    Next
    If somedeleted > 0 Then GoTo lIteration
End Sub

'Check file len safely
Public Function FileLen_safe(sFileName As String) As Long
    On Error Resume Next
    FileLen_safe = -1
    FileLen_safe = FileLen(sFileName)
End Function


'Root function of optimization
'Main generic version for optimizing highways and junctions
Public Sub OptimizeRouting(InputFile As String)
    Dim OutFile As String
    'Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'no file - nothing to do
    If FileLen_safe(InputFile) < 1 Then Exit Sub 'empty or missing file
    
    OutFile = InputFile + "_opt.mp" 'output file
    'OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'Join nodes by NodeID
    Call JoinNodesByID
    DoEvents
    
    'Join two way roads into bidirectional ways
    Call JoinDirections3(70, -0.996, -0.95, 100, 2)
    '70 metres between directions (Ex: Universitetskii pr, Moscow - 68m)
    '-0.996 -> (175, 180) degrees for start contradirectional check
    '-0.95 -> (161.8, 180) degrees for further contradirectional checks
    '100 metres min two way road
    '2 metres for joining nodes into one
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Call Save_MP(OutFile2)  'temp file for debug
    'DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm with limiting edge len
    Call DouglasPeucker_total_split(5, 100)
    'Epsilon = 5 metres
    'Max edge - 100 metres
    DoEvents
    
    Call CollapseJunctions2(1000, 1200, 0.13)
    'Slide allowed up to 1000 metres
    'Max junction loop is 1200 metres
    '0.13 -> ~ 7.46 degress
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5)
    'Epsilon = 5 metres
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5)
    'Epsilon = 5 metres
    DoEvents
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(3)
    'CollapseDistance = 3 metres
    DoEvents
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub

'##########################################################################################
'
'Block of code, used only by additional optimization functions
'for Planet Overview and so on

'Generalize highways for Planet Overview
Public Sub OptimizeRouting_hw(InputFile As String) '_hw
    Dim OutFile As String
    'Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    'OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.05   '0.05 degrees for local maps, 1 for planet-s
    Control_ForceWaySpeed = 4    'set -1 to not force, 0 or more to forcing this value
    Control_TrunkType = 2        'set 1 to be have same as motorway = 0x01 Major highway
    Control_PrimaryType = 3      'set 2 to use 0x02 Principal highway
    Control_TrunkLinkType = 8    'set 9 to have same as motorway
    Control_LoadNoRoute = 0      'set 0 to skip no-routing polylines, 1 to load
    Control_LoadMPType = 0       'set 0 to skip mp Type= field, 1 to parse
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'Join nodes by NodeID
    Call JoinNodesByID
    DoEvents
    
    'Join two way roads into bidirectional ways
    Call JoinDirections3(70, -0.996, -0.95, 100, 2)
    '70 metres between directions (Ex: Universitetskii pr, Moscow - 68m)
    '-0.996 -> (175, 180) degrees for start contradirectional check
    '-0.95 -> (161.8, 180) degrees for further contradirectional checks
    '100 metres min two way road
    '2 metres for joining nodes into one
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Call Save_MP(OutFile2)  'temp file for debug
    'DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm with limiting edge len
    Call DouglasPeucker_total_split(5, 100)
    'Epsilon = 5 metres
    'Max edge - 100 metres
    DoEvents
    
    Call CollapseJunctions2(3000, 7000, 0.13)
    'Slide allowed up to 3000 metres
    'Max junction loop is 6000 metres
    '0.13 -> ~ 7.46 degress
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call RemoveOneWay
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5)
    'Epsilon = 5 metres
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(500)
    'Epsilon = 500 metres
    DoEvents
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(300)
    'CollapseDistance = 300 metres
    DoEvents
    
    'Combine close nodes and remove duplicate edges
    Call JoinCloseNodes(200) '200 metres
    DoEvents
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'Generalize highways for Planet Overview with trim by bbox
'param of launch: filename?a?b?c?d    a - lon min, b - lat min, c - lon max, d - lat max
Public Sub OptimizeRouting_hwbbox(Cmd As String) '_hwbbox
    Dim OutFile As String
    Dim CmdArgs() As String
    Dim InputFile As String
    Dim bbox1 As bbox
    Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    CmdArgs = Split(Cmd, "?")
    If UBound(CmdArgs) < 0 Then Exit Sub 'nothing to do
    
    InputFile = CmdArgs(0)
    If InputFile = "" Then Exit Sub 'no file - nothing to do
    If FileLen_safe(InputFile) < 1 Then Exit Sub 'empty or missing file
    
    bbox1.lat_min = -360
    If UBound(CmdArgs) >= 4 Then
        bbox1.lon_min = Val(CmdArgs(1))
        bbox1.lat_min = Val(CmdArgs(2))
        bbox1.lon_max = Val(CmdArgs(3))
        bbox1.lat_max = Val(CmdArgs(4))
    End If
    
    OutFile = InputFile + "_opt.mp" 'output file
    'OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.05   '0.05 degrees for local maps, 1 for planet-s
    Control_ForceWaySpeed = 4    'set -1 to not force, 0 or more to forcing this value
    Control_TrunkType = 2        'set 1 to be have same as motorway = 0x01 Major highway
    Control_PrimaryType = 3      'set 2 to use 0x02 Principal highway
    Control_TrunkLinkType = 8    'set 9 to have same as motorway
    Control_LoadNoRoute = 0      'set 0 to skip no-routing polylines, 1 to load
    Control_LoadMPType = 0       'set 0 to skip mp Type= field, 1 to parse
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'Join nodes by NodeID
    Call JoinNodesByID
    DoEvents
    
    'Join two way roads into bidirectional ways
    Call JoinDirections3(70, -0.996, -0.95, 100, 2)
    '70 metres between directions (Ex: Universitetskii pr, Moscow - 68m)
    '-0.996 -> (175, 180) degrees for start contradirectional check
    '-0.95 -> (161.8, 180) degrees for further contradirectional checks
    '100 metres min two way road
    '2 metres for joining nodes into one
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Call Save_MP(OutFile2)  'temp file for debug
    'DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm with limiting edge len
    Call DouglasPeucker_total_split(5, 100)
    'Epsilon = 5 metres
    'Max edge - 100 metres
    DoEvents
    
    Call CollapseJunctions2(3000, 7000, 0.13)
    'Slide allowed up to 3000 metres
    'Max junction loop is 7000 metres
    '0.13 -> ~ 7.46 degress
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call RemoveOneWay
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5)
    'Epsilon = 5 metres
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(500)
    'Epsilon = 500 metres
    DoEvents
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(300)
    'CollapseDistance = 300 metres
    DoEvents
    
    'Combine close nodes and remove duplicate edges
    Call JoinCloseNodes(200) '200 metres
    DoEvents
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Trim all data by bbox
    If bbox1.lat_min > -360 Then
        Call TrimByBbox(bbox1)
    End If
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'Generalize borders (all levels)
Public Sub OptimizeRouting_borders(InputFile As String) '_borders
    Dim OutFile As String
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.1
    Control_ForceWaySpeed = 4
    Control_TrunkType = 2
    Control_TrunkLinkType = 8
    Control_LoadNoRoute = 1
    Control_LoadMPType = 1
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    Call JoinCloseNodes(1) '1 metre
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(500) 'for zooms 0-2
    'Call DouglasPeucker_total(5000) 'for zoom 3 e.t.c.
    DoEvents
    
    Call CollapseShortEdges(10)
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub

'Generalize borders for top levels
Public Sub OptimizeRouting_borders_top(InputFile As String) '_borders_top
    Dim OutFile As String
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.1
    Control_ForceWaySpeed = 4
    Control_TrunkType = 2
    Control_TrunkLinkType = 8
    Control_LoadNoRoute = 1
    Control_LoadMPType = 1
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    Call JoinCloseNodes(1) '1 metre
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5000) 'for zoom 3 e.t.c.
    DoEvents
    
    Call CollapseShortEdges(10)
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'Generalize railroads for Planet Overview
Public Sub OptimizeRouting_rr(InputFile As String) '_rr
    Dim OutFile As String
    Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.1
    Control_ForceWaySpeed = 4   'set -1 to not force
    Control_TrunkType = 2      'set 1 to be have same as motorway
    Control_TrunkLinkType = 8 'set 9 to have same as motorway
    Control_LoadNoRoute = 1   'set 0 to skip no-routing polylines
    Control_LoadMPType = 1        'set 0 to skip mp Type field
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    Call JoinCloseNodes(30) '100 m
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call CollapseShortEdges(30)
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    Call CollapseShortEdges(20)
    DoEvents
    
    Call CollapseJunctions2(3000, 7000, 0.13)
    'Slide allowed up to 3000 metres
    'Max junction loop is 7000 metres
    '0.13 -> ~ 7.46 degress
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(100) 'for zooms 0-
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(100)
    DoEvents
    
    Call JoinAcute(100, 3)
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
lSkip1:
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'Generalize railroads for Planet Overview with trim by bbox
'param of launch: filename?a?b?c?d    a - lon min, b - lat min, c - lon max, d - lat max
Public Sub OptimizeRouting_rrbbox(Cmd As String) 'rrbbox
    Dim OutFile As String
    Dim CmdArgs() As String
    Dim InputFile As String
    Dim bbox1 As bbox
    Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    CmdArgs = Split(Cmd, "?")
    If UBound(CmdArgs) < 0 Then Exit Sub 'nothing to do
    
    InputFile = CmdArgs(0)
    If InputFile = "" Then Exit Sub 'no file - nothing to do
    If FileLen_safe(InputFile) < 1 Then Exit Sub 'empty or missing file
    
    bbox1.lat_min = -360
    If UBound(CmdArgs) >= 4 Then
        bbox1.lon_min = Val(CmdArgs(1))
        bbox1.lat_min = Val(CmdArgs(2))
        bbox1.lon_max = Val(CmdArgs(3))
        bbox1.lat_max = Val(CmdArgs(4))
    End If
    
    OutFile = InputFile + "_opt.mp" 'output file
    OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.1
    Control_ForceWaySpeed = 4   'set -1 to not force
    Control_TrunkType = 2      'set 1 to be have same as motorway
    Control_TrunkLinkType = 8 'set 9 to have same as motorway
    Control_LoadNoRoute = 1   'set 0 to skip no-routing polylines
    Control_LoadMPType = 1        'set 0 to skip mp Type field
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    Call JoinCloseNodes(30) '30 m
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    '#Call CollapseShortEdges(30)
    Call CollapseShortEdges(30)
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    Call CollapseShortEdges(20)
    DoEvents
    
    Call CollapseJunctions2(3000, 7000, 0.13)
    'Slide allowed up to 1000 metres
    'Max junction loop is 1200 metres
    '0.13 -> ~ 7.46 degress
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(100) 'for zooms 0-
    'Epsilon = 100 m
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(100)
    DoEvents
    
    Call JoinAcute(100, 3)
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    'Trim all data by bbox, if specified
    If bbox1.lat_min > -360 Then
        Call TrimByBbox(bbox1)
    End If
    
lSkip1:
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'For combining railroads or highways for Planet Overview
'old way, without bbox trim/stitch
Public Sub OptimizeRouting_comb(InputFile As String) '_comb
    Dim OutFile As String
    Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    Control_ClusterSize = 0.1    '0.05 degrees for local maps, 1 for planet-s
    Control_ForceWaySpeed = 4    'set -1 to not force, 0 or more to forcing this value
    Control_TrunkType = 2        'set 1 to be have same as motorway = 0x01 Major highway
    Control_PrimaryType = 3      'set 2 to use 0x02 Principal highway
    Control_TrunkLinkType = 8    'set 9 to have same as motorway
    Control_LoadNoRoute = 1      'set 0 to skip no-routing polylines, 1 to load
    Control_LoadMPType = 1       'set 0 to skip mp Type= field, 1 to parse
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'Call Save_MP(OutFile2)
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    Call JoinCloseNodes(100) '100 metres
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call CollapseShortEdges(30)
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Join edges with very acute angle into one
    Call JoinAcute(100, 3)
    '100 metres for joining nodes
    'AcuteKoeff = 3 => 18.4 degrees
    DoEvents
    
    Call CollapseShortEdges(20)
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(100) 'for zooms 0+
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(100)
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
lSkip1:
    
    'Save result
    'Call Save_MP(OutFile2)
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'For combining highways or railways for Planet Overview with stitching on 5x5 and 1x1 borders
Public Sub OptimizeRouting_stitch(InputFile As String) '_stitch
    Dim OutFile As String
    'Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    Dim bbox1 As bbox
    Dim i As Long, j As Long
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    'OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.1    '0.05 degrees for local maps, 1 for planet-s
    Control_ForceWaySpeed = 4    'set -1 to not force, 0 or more to forcing this value
    Control_TrunkType = 2        'set 1 to be have same as motorway = 0x01 Major highway
    Control_PrimaryType = 3      'set 2 to use 0x02 Principal highway
    Control_TrunkLinkType = 8    'set 9 to have same as motorway
    Control_LoadNoRoute = 0      'set 0 to skip no-routing polylines, 1 to load
    Control_LoadMPType = 1       'set 0 to skip mp Type= field, 1 to parse
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call JoinCloseNodes(1) 'join close nodes, not by ID as file expected to be concat of different files with id collisions
    DoEvents
    
    'index nodes
    Call BuildNodeClusterIndex(0)
    
    '1. Stitch all world-wide lines
    For i = -175 To 175 Step 5
        bbox1.lat_min = -90
        bbox1.lat_max = 90
        bbox1.lon_min = i
        bbox1.lon_max = i
        'beware, double type have finite precision
        'recommended to setup integer coordinates
        'or having N/2^m part (.5, .25, .375, .5625 and so on)
        'or range, ex. min=29.39999,max=29.40001
        
        Call StitchNodes(bbox1, 1000) '1000m
        DoEvents
    Next
    
    For i = -85 To 85 Step 5
        bbox1.lat_min = i
        bbox1.lat_max = i
        bbox1.lon_min = -180
        bbox1.lon_max = 180
        
        Call StitchNodes(bbox1, 1000) '1000m
        DoEvents
    Next
    
    '2. Stitch all small lines
    
    'example for highways
    Call Stitch1x1from5x5(40, -75, 1000)
    Call Stitch1x1from5x5(35, 135, 1000)
    Call Stitch1x1from5x5(50, -5, 1000)
    Call Stitch1x1from5x5(50, 5, 1000)
    Call Stitch1x1from5x5(45, 0, 1000)
    Call Stitch1x1from5x5(45, 5, 1000)
    Call Stitch1x1from5x5(45, 10, 1000)

    'example for railways
'    For j = 45 To 50 Step 5
'    For i = -5 To 10 Step 5
'    Call Stitch1x1from5x5(CDbl(j), CDbl(i), 1000)
'    Next
'    Next

    'Join close nodes afterward, just in case
    Call JoinCloseNodes(1)
    DoEvents
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'For sorting peaks and volcanos to zoom levels
'Load and save osm file(s), not mp
Public Sub OptimizeRouting_ele(InputFile As String) '_ele
    Dim OutFile As String
    'Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_lev" 'output file
    'OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.5
    'other control_ are irrelevant
    
    'Fake world bbox - by two nodes
    Nodes(NodesNum).lat = -90
    Nodes(NodesNum).lon = -180
    AddNode
    Nodes(NodesNum).lat = 90
    Nodes(NodesNum).lon = 180
    AddNode
    Call BuildNodeClusterIndex(0)
    
    Nodes(0).NodeID = MARK_NODEID_DELETED 'mark two fake nodes as deleted
    Nodes(1).NodeID = MARK_NODEID_DELETED
    
    'Load data from file
    Call Load_OSM_lined(InputFile)
    DoEvents
    
    'Sort nodes
    Call SortNodesByEle
    DoEvents
    
    'Save to 4 out files
    Call Save_OSM_lined(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub

'For optimizing bathymetry from NaturalEarth
Public Sub OptimizeRouting_bathy(InputFile As String) '_bathy
    Dim OutFile As String
    Dim OutFile2 As String 'temp file for debug
    Dim time1 As Double
    
    If InputFile = "" Then Exit Sub 'nothing to do
    
    OutFile = InputFile + "_opt.mp" 'output file
    OutFile2 = InputFile + "_p.mp"  'output2 - for intermediate results
    
    time1 = Timer 'start measure time
    
    'Init module (all arrays)
    Call init
    
    Control_ClusterSize = 0.5
    Control_ForceWaySpeed = 4
    Control_TrunkType = 2
    Control_TrunkLinkType = 8
    Control_LoadNoRoute = 1
    Control_LoadMPType = 1
    
    'Load data from file
    Call Load_MP(InputFile, 1200)
    DoEvents
    
    'Call Save_MP(OutFile2)
    
    'No data - do nothing
    If NodesNum < 1 And EdgesNum < 1 Then Exit Sub
    
    Call RemoveOneWay
    DoEvents
    
    'Call JoinCloseNodes(15) '15 m - for ne_10m_bathymetry_K_200.mp
    Call JoinCloseNodes(1) '1 m - for all other
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call CollapseShortEdges(500)
    'DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
    Call FilterVoidEdges
    DoEvents
    
    'Optimize all roads by (Ramer–)Douglas–Peucker algorithm
    Call DouglasPeucker_total(5000) 'for zooms 0-
    
    
    'Remove very short edges, they are errors, most probably
    Call CollapseShortEdges(1000)
    DoEvents
    
    Call CombineDuplicateEdgesAll
    DoEvents
    
lSkip1:
    
    'Save result
    Call Save_MP_2(OutFile)
    
    Form1.Caption = "Done " + Format(Timer - time1, "0.00") + " s" 'display timing

End Sub


'Remove flag oneway from all edges
Public Sub RemoveOneWay()
    Dim i As Long
    For i = 0 To EdgesNum - 1
        Edges(i).oneway = 0
        If (i And 32767) = 0 Then
            'show progress
            Form1.Caption = "RemoveOneWay " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    Next
End Sub


'Combinde pair of edges, which connects same nodes (including specified one)
Public Sub CombineDuplicateEdges(ByVal node1 As Long)
    Dim i As Long, j As Long, e1 As Long, e2 As Long
    Dim node2 As Long
    
    i = 0
    
    While i < Nodes(node1).Edges - 1
        e1 = Nodes(node1).edge(i)
        node2 = Edges(e1).node2
        If node2 = node1 Then node2 = Edges(e1).node1
        
        j = i + 1
        While j < Nodes(node1).Edges
            e2 = Nodes(node1).edge(j)
            If Edges(e2).node1 = node2 Or Edges(e2).node2 = node2 Then
                'other end nodes are the same - combine
                Call CombineEdges(e1, e2, node1)
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
End Sub


'Combinde pair of edges, which connects same nodes
Public Sub CombineDuplicateEdgesAll()
    Dim i As Long
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED Then
            Call CombineDuplicateEdges(i) 'check edges of each node
        End If
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "CombineDuplicateEdgesAll " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
    Next
End Sub


'Join nodes close to each other
'DistLimit - max distance for joining
Public Sub JoinCloseNodes(ByVal DistLimit As Double)
    Dim i As Long, j As Long
    Dim mode1 As Long
    Dim bbox1 As bbox
    Dim DistLimitSq As Double
    
    'build index for finding close nodes
    DistLimitSq = DistLimit * DistLimit
    Call BuildNodeClusterIndex(0)
    
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID <> MARK_NODEID_DELETED Then
            
            'Create bbox for area around node
            mode1 = 0
            bbox1.lat_max = Nodes(i).lat
            bbox1.lat_min = Nodes(i).lat
            bbox1.lon_max = Nodes(i).lon
            bbox1.lon_min = Nodes(i).lon
            Call ExpandBbox(bbox1, DistLimit)
            
lNextNode:
            j = GetNodeInBboxByCluster(bbox1, mode1)
            mode1 = 1 '"next" next time
            If j = -1 Then GoTo lAllNodes 'no more nodes
            
            If i <> j And Nodes(j).NodeID <> MARK_NODEID_DELETED And DistanceSquare(i, j) < DistLimitSq Then
                'skip same node, deleted nodes and too far
                
                Call MergeNodes(i, j, 1) 'merge
                Call CombineDuplicateEdges(i) 'removed duplicate edges
            End If
            GoTo lNextNode
lAllNodes:
    
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "JoinCloseNodes (" + CStr(DistLimit) + ") " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
        End If
            
        End If
    Next


End Sub


'Sort nodes to zoom levels by location and elevation
'Assumes:
'1) that nodes already sorted be elevation - first nodes must have max elevation
'2) value of parsed tag "ele" must be in .temp_dist field
'3) ClusterIndex is initialized, but no filled with (live) nodes
'Function sets .mark field to 1, 2 and 3 to indicate zoom
'To zoom 3 - all nodes are not close than 150km
'To zoom 2 - all nodes are not close than 40km
'To zoom 1 - all nodes are not close than 10km
'Rest nodes remain on zoom 0
'Nodes with unclear ele and <1000 - always on zoom 0
Public Sub SortNodesByEle()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim DistLimit As Double
    Dim DistLimitSq As Double
    Dim bbox1 As bbox
    Dim mode1 As Long
    Dim MinEle As Double
    
    MinEle = 1000 'below - to zoom 0
    
    ReDim Preserve ClustersChain(NodesNum) 'allocate for whole chain
    
    For k = 3 To 1 Step -1
        Select Case k
            Case 1 'zooms 2.1-8km
                DistLimit = 10000 '10km
            Case 2 'zooms 8.1-30km
                DistLimit = 40000 '40km
            Case 3 'zooms 31-120 km
                DistLimit = 150000 '150km
            End Select
            DistLimitSq = DistLimit * DistLimit
        
        For i = 0 To NodesNum - 1
            If Nodes(i).mark > 0 Then GoTo lSkipNode 'already higher level
            If Nodes(i).temp_dist < MinEle Then GoTo lSkipNode 'too small
            If Nodes(i).NodeID = MARK_NODEID_DELETED Then GoTo lSkipNode 'deleted
            
            'Create bbox around node
            mode1 = 0
            bbox1.lat_max = Nodes(i).lat
            bbox1.lat_min = Nodes(i).lat
            bbox1.lon_max = Nodes(i).lon
            bbox1.lon_min = Nodes(i).lon
            Call ExpandBbox(bbox1, DistLimit)
            
            'Search to indexed nodes around this one
lNextGet:
            j = GetNodeInBboxByCluster(bbox1, mode1)
            mode1 = 1
            If j = -1 Then GoTo lNotFound
            
            If DistanceSquare(i, j) < DistLimitSq Then GoTo lSkipNode
            GoTo lNextGet
lNotFound:
            'no indexed nodes found - add this one to index
            Nodes(i).mark = k 'store found zoom level
            Call AddNodeToClusterIndex(i)
lSkipNode:
        
            If (i And 8191) = 0 Then
                'show progress
                Form1.Caption = "SortNodesByEle level " + CStr(k) + ", " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
            End If
        
        Next
    Next

End Sub


'Loosely load nodes from osm file
'Expect for osm file to have all node info in one line
'Does not parse XML, only found all need info
'field .label in Edge used to store lines of OSM XML for further resave
Public Sub Load_OSM_lined(filename As String)
    Dim i As Long, j As Long, k As Long
    Dim FileLen As Long
    Dim sLine As String
    Dim fLat As Double
    Dim fLon As Double
    Dim ele_str As String
    Dim ele As Double
    Dim NodeID As String
    
    NodeIDMax = -1 'no nodeid yet
    
    Open filename For Input As #1
    FileLen = LOF(1)
    
lNextLine:
    Line Input #1, sLine
    
    'check POI type
    i = InStr(1, sLine, "<tag k=""natural"" v=""peak""/>")
    j = InStr(1, sLine, "<tag k=""natural"" v=""volcano""/>")
    If i < 1 And j < 1 Then GoTo lSkipNode
    
    'check presence of node id
    i = InStr(1, sLine, "<node id=""")
    If i < 1 Then GoTo lSkipNode 'no id
    j = InStr(i + 10, sLine, """")
    If j < 1 Then GoTo lSkipNode 'no ending "
    
    NodeID = Mid(sLine, i + 10, j - i - 10)
    
    'parse lat
    i = InStr(1, sLine, " lat=""")
    If i < 1 Then GoTo lSkipNode 'no lat - skip line
    
    fLat = Val(Mid(sLine, i + 6, 20))
    
    'parse lon
    i = InStr(1, sLine, " lon=""")
    If i < 1 Then GoTo lSkipNode
    
    fLon = Val(Mid(sLine, i + 6, 20))
    
    'parse ele
    i = InStr(1, sLine, " k=""ele""")
    If i < 1 Then GoTo lSkipNode 'no ele
    
    j = InStr(i + 12, sLine, """")
    If j < 1 Then GoTo lSkipNode 'no ending "
    ele_str = Mid(sLine, i + 12, j - i - 12)
    ele = Val(ele_str)
    If InStr(1, ele_str, "f") > 0 Then
        'f - feet, elevation in feets will be treated as 0
        ele = 0
    End If
    
    'save node info
    j = NodesNum
    Nodes(NodesNum).lat = fLat
    Nodes(NodesNum).lon = fLon
    Nodes(NodesNum).Edges = 0
    Nodes(NodesNum).NodeID = -1
    Nodes(NodesNum).temp_dist = ele
    Nodes(NodesNum).mark = 0
    Call AddNode
    
    'save edge info - label = line of OSM
    k = EdgesNum
    Edges(EdgesNum).node1 = j
    Edges(EdgesNum).node2 = j
    'Edges(EdgesNum).label = NodeID
    Edges(EdgesNum).label = sLine
    Call AddEdge
    
    Call AddEdgeToNode(j, k)
    
    
    If (NodesNum And 1023) = 0 Then
        'display progress
        Form1.Caption = "Load_OSM_lined: " + CStr(Seek(1)) + " / " + CStr(FileLen): Form1.Refresh
    End If
    

lSkipNode:
    If Not EOF(1) Then GoTo lNextLine
    
    Close #1

End Sub


'Save OSM files from similar loaded earlier (see Load_OSMlined)
'Saved 4 files - each for 4 zooms from 0 to 3
Public Sub Save_OSM_lined(filename As String)
    Dim i As Long, j As Long
    Dim k1 As Long, k2 As Long
    Dim typ As Long
    
    For k1 = 0 To 3
        Open filename + CStr(k1) + ".osm" For Output As #2
        'OSM header
        Print #2, "<?xml version='1.0' encoding='UTF-8'?>"
        Print #2, "<osm version='0.6' generator='mp_extsimp1'>"
        
        'save lines from nodes of specific zoom
        For i = 0 To NodesNum - 1
            If Nodes(i).NodeID = MARK_NODEID_DELETED Then GoTo lSkipNode
            If Nodes(i).mark <> k1 Then GoTo lSkipNode
            j = Nodes(i).edge(0)
            Print #2, Edges(j).label
            
            If (i And 8191) = 0 Then
                'display progress
                Form1.Caption = "Save_OSM_lined " + CStr(i) + " / " + CStr(NodesNum): Form1.Refresh
            End If
            
lSkipNode:
        Next
        Print #2, "</osm>" 'ending marker
        Close #2
        
    Next
End Sub


'Parse mp Type to our own constants
Public Function GetTypeFromMP(MPType As Long) As Long
    Select Case MPType
    Case &H1 'Major highway
        GetTypeFromMP = HIGHWAY_MOTORWAY
    Case &H2 'Principal highway
        GetTypeFromMP = HIGHWAY_TRUNK
    Case &H3 'Other highway road
        GetTypeFromMP = HIGHWAY_PRIMARY
        
    'block for simplification of borders
    Case &H1E 'Country border
        GetTypeFromMP = HIGHWAY_PRIMARY
    Case &H1C 'State/region border
        GetTypeFromMP = HIGHWAY_SECONDARY
    Case Else
        GetTypeFromMP = HIGHWAY_SECONDARY
    End Select
End Function


'Add 1 node to ClusterIndex array (array must be prepared)
'Warning: should not be called before BuildNodeClusterIndex(1)
Public Sub AddNodeToClusterIndex(node1 As Long)
    Dim i As Long, j As Long, k As Long
    Dim x As Long
    Dim y As Long
    
    i = node1
    If Nodes(i).NodeID <> MARK_NODEID_DELETED Then
        'get cluster from lat/lon
        x = (Nodes(i).lat - ClustersLat0) / Control_ClusterSize
        y = (Nodes(i).lon - ClustersLon0) / Control_ClusterSize
        j = x + y * ClustersLatNum
        
        k = ClustersLast(j)
        If k = -1 Then
            'first index in chain of this cluster
            ClustersFirst(j) = i
        Else
            'continuing chain
            ClustersChain(k) = i
        End If
        ClustersChain(i) = -1 'this is last node in chain
        ClustersLast(j) = i
    End If
    ClustersIndexedNodes = 0 ' fake value, as node1 may not be last in indexed

End Sub


'Trim map by specified bbox
'All edges crossing bbox will be cut by inserting new node and reconnecting edge from outside node to new one
'All nodes outside bbox will be deleted
'Should correctly handle edges on corner and bbox-wide edges
Public Sub TrimByBbox(bbox1 As bbox)
    Dim i As Long, j As Long, k As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim px As Long
    
    'Trim all edges crossing bbox
    For i = 0 To EdgesNum - 1
    If Edges(i).node1 = -1 Then GoTo lSkip1
        
lCheckAgain:
        p1 = Edges(i).node1
        p2 = Edges(i).node2
        
        If Nodes(p1).lat < bbox1.lat_min And Nodes(p2).lat > bbox1.lat_min Then
            'trim edge by lat_min
            px = p1 'node1 will be deleted
            GoTo lTrim1
        End If
        If Nodes(p2).lat < bbox1.lat_min And Nodes(p1).lat > bbox1.lat_min Then
            'trim edge by lat_min
            px = p2 'node2 will be deleted
lTrim1:
            j = NodesNum
            Nodes(j).lat = bbox1.lat_min 'new node coordinates
            Nodes(j).lon = Nodes(p1).lon + (Nodes(p2).lon - Nodes(p1).lon) * (bbox1.lat_min - Nodes(p1).lat) / (Nodes(p2).lat - Nodes(p1).lat)
            Call AddNode
            Call ReconnectEdge(i, px, j)
            GoTo lCheckAgain  'check edge again as it may cross other sides of bbox as well
        End If

        If Nodes(p1).lat < bbox1.lat_max And Nodes(p2).lat > bbox1.lat_max Then
            'trim edge by lat_max
            px = p2 'node2 will be deleted
            GoTo lTrim2
        End If
        If Nodes(p2).lat < bbox1.lat_max And Nodes(p1).lat > bbox1.lat_max Then
            'trim edge by lat_min
            px = p1 'node1 will be deleted
lTrim2:
            j = NodesNum
            Nodes(j).lat = bbox1.lat_max
            Nodes(j).lon = Nodes(p1).lon + (Nodes(p2).lon - Nodes(p1).lon) * (bbox1.lat_max - Nodes(p1).lat) / (Nodes(p2).lat - Nodes(p1).lat)
            Call AddNode
            Call ReconnectEdge(i, px, j)
            GoTo lCheckAgain
        End If

        If Nodes(p1).lon < bbox1.lon_min And Nodes(p2).lon > bbox1.lon_min Then
            'trim edge by lon_min
            px = p1 'node1 will be deleted
            GoTo lTrim3
        End If
        If Nodes(p2).lon < bbox1.lon_min And Nodes(p1).lon > bbox1.lon_min Then
            'trim edge by lon_min
            px = p2 'node2 will be deleted
lTrim3:
            j = NodesNum
            Nodes(j).lat = Nodes(p1).lat + (Nodes(p2).lat - Nodes(p1).lat) * (bbox1.lon_min - Nodes(p1).lon) / (Nodes(p2).lon - Nodes(p1).lon)
            Nodes(j).lon = bbox1.lon_min
            Call AddNode
            Call ReconnectEdge(i, px, j)
            GoTo lCheckAgain
        End If

        If Nodes(p1).lon < bbox1.lon_max And Nodes(p2).lon > bbox1.lon_max Then
            'trim edge by lon_max
            px = p2 'node2 will be deleted
            GoTo lTrim4
        End If
        If Nodes(p2).lon < bbox1.lon_max And Nodes(p1).lon > bbox1.lon_max Then
            'trim edge by lon_max
            px = p1 'node1 will be deleted
lTrim4:
            j = NodesNum
            Nodes(j).lat = Nodes(p1).lat + (Nodes(p2).lat - Nodes(p1).lat) * (bbox1.lon_max - Nodes(p1).lon) / (Nodes(p2).lon - Nodes(p1).lon)
            Nodes(j).lon = bbox1.lon_max
            Call AddNode
            Call ReconnectEdge(i, px, j)
            GoTo lCheckAgain
        End If

lSkip1:
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "TrimByBbox, trim " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If
    Next
    
    'Now no single edge cross bbox
    
    'Delete all nodes outside bbox with remaining edges
    For i = 0 To NodesNum - 1
        If Nodes(i).NodeID = MARK_NODEID_DELETED Then GoTo lSkip2
    
        If Nodes(i).lat < bbox1.lat_min Or _
            Nodes(i).lat > bbox1.lat_max Or _
            Nodes(i).lon < bbox1.lon_min Or _
            Nodes(i).lon > bbox1.lon_max Then
            Call DelNode(i)
        End If

lSkip2:
        If (i And 8191) = 0 Then
            'show progress
            Form1.Caption = "TrimByBbox, del " + CStr(i) + " / " + CStr(EdgesNum): Form1.Refresh
        End If

    Next
End Sub


'Function to stitch roads, trimmed by two near bbox
'bbox1 - bbox of/around border
'MaxDist - max distance between nodes to stitch
'Assumed, that clusterindex were built before the call
Public Sub StitchNodes(bbox1 As bbox, MaxDist As Double)
    Dim i As Long, k As Long
    Dim j As Long, k2 As Long
    Dim p As Long
    Dim mode1 As Long
    Dim bbox2 As bbox
    Dim d As Double
    Dim edge1 As Long, edge2 As Long
    Dim MaxDistSQ As Double
    Dim edge_cos As Double
    Dim bbox_type As Long '0 - same lat, 1 - same lon
    Dim MinDist As Double, nodeMinDist As Long
    
    MaxDistSQ = MaxDist * MaxDist
    
    'check bbox type
    If (bbox1.lat_max - bbox1.lat_min) < (bbox1.lon_max - bbox1.lon_min) Then
        '  same lat: ---
        bbox_type = 0
    Else
        '  same lon: |
        bbox_type = 1
    End If

    'show progress
    Form1.Caption = "StitchNodes (" + CStr(bbox1.lat_min) + "," + CStr(bbox1.lon_min) + ";" + CStr(MaxDist) + ")": Form1.Refresh
    
    'index all nodes on border to Chain
    ChainNum = 0
    
    mode1 = 0
lNextNode:
    j = GetNodeInBboxByCluster(bbox1, mode1)
    mode1 = 1 '"next" next time
    If j > -1 Then
    
        If Nodes(j).Edges <> 1 Then GoTo lNextNode 'skip not ends
        Call AddChain(j)
        Nodes(j).mark = 0 'edge to the botton (south) or left (west)
        
        k = Nodes(j).edge(0)
        i = Edges(k).node1
        If i = j Then i = Edges(k).node2
        If bbox_type = 0 Then
            If Nodes(i).lat > bbox1.lat_max Then Nodes(j).mark = 1 ' edge to the top (north)
        Else
            If Nodes(i).lon > bbox1.lon_max Then Nodes(j).mark = 1 ' edge to the right (east)
        End If
        
        GoTo lNextNode
    End If
    
    
    'check all indexed node, searching best for joining
    
    For i = 0 To ChainNum - 1
        k = Chain(i)
        If Nodes(k).NodeID = MARK_NODEID_DELETED Or Nodes(k).Edges <> 1 Then GoTo lSkip1 'skip already stitched
        
        nodeMinDist = -1
        MinDist = MaxDistSQ
        
        For j = i + 1 To ChainNum - 1
            k2 = Chain(j)
            If Nodes(k2).mark = Nodes(k).mark Then GoTo lSkip2 'edges at the same side of stitch line
            If Nodes(k2).NodeID = MARK_NODEID_DELETED Then GoTo lSkip2 'skip already stitched
            If Nodes(k2).Edges <> 1 Then GoTo lSkip2 'skip already stitched
            d = DistanceSquare(k, k2)
            If d > MaxDistSQ Then GoTo lSkip2 'node too far
            
            edge1 = Nodes(k).edge(0)
            edge2 = Nodes(k2).edge(0)
            If Edges(edge1).roadtype <> Edges(edge2).roadtype Then GoTo lSkip2 'do not stitch different road classes
            'speed class is ignored for now
            
            edge_cos = CosAngleBetweenEdges(edge1, edge2) 'calc cosine between edges
            
            d = d * (1 - Abs(edge_cos)) 'weight sq-distance and angle
            'the lesser distance or angle between edges - the better
            
            'find best matched node for selected one (k)
            If d < MinDist Then MinDist = d: nodeMinDist = k2
            
lSkip2:
        Next
lSkip1:
    
        'if some node found - stitch it
        If nodeMinDist > -1 Then
            Call MergeNodes(k, nodeMinDist)
        End If
    
        If (i And 255) = 0 Then
            'show progress
            Form1.Caption = "StitchNodes (" + CStr(bbox1.lat_min) + "," + CStr(bbox1.lon_min) + ";" + CStr(MaxDist) + ") " + CStr(i) + " / " + CStr(ChainNum): Form1.Refresh
        End If
    Next

    
End Sub

'stitch internal borders between 1x1 degrees bboxes inside 5x5 degrees bbox
Public Sub Stitch1x1from5x5(lat As Double, lon As Double, MaxDist As Double)
    Dim bbox1 As bbox
    Dim i As Long
    For i = 1 To 4
        bbox1.lat_min = lat + i
        bbox1.lat_max = lat + i
        bbox1.lon_min = lon
        bbox1.lon_max = lon + 5
        Call StitchNodes(bbox1, MaxDist)
    Next

    For i = 1 To 4
        bbox1.lat_min = lat
        bbox1.lat_max = lat + 5
        bbox1.lon_min = lon + i
        bbox1.lon_max = lon + i
        Call StitchNodes(bbox1, MaxDist)
    Next

End Sub
