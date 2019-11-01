// Variable that is used to store the current scroll position of a window
var yoffset = 0

// The Tree object constructor takes the name of a variable
// to which the tree object is assigned (for use in event 
// handling) and a unique id value (to differentiate one tree
// from another)
function Tree(variable, id) {
 this.variable = variable		// The name of a variable by which the tree can be referenced
 this.id = id				// A unique ID for the tree
 this.nodeID = 0			// Used to give each node of the tree a unique ID
 this.nodes = new Array()		// An array consisting of the root nodes of the tree
 this.openedMarker = "-  "	// String used to indicate that a node has been opened
 this.closedMarker = "+  "	// String used to indicate that a node has been closed
 this.leafMarker = "*  "	// String used to indicate a leaf of the tree
 this.addNode = Tree_addNode		// Add a root node to the tree
 this.display = Tree_display		// Display the tree
 this.loadState = Tree_loadState	// Load the state of the tree from cookies
 this.saveState = Tree_saveState	// Save the state of the tree using cookies
 this.getState = Tree_getState		// Get a string that contains the current state of each tree node
 this.setState = Tree_setState		// Set the state of each tree node from a string
 this.getNodeID = Tree_getNodeID	// Returns a new node ID
 this.getNode = Tree_getNode		// Returns the node corresponding to a specific ID
 this.setMarkers = Tree_setMarkers	// Sets the opened, closed, and leaf markers of the tree
}

// Add a root node to the tree
function Tree_addNode(node) {
 this.nodes[this.nodes.length] = node
}

// Display the tree object
function Tree_display() {
 for(var i=0; i<this.nodes.length; ++i) 
  this.nodes[i].display()
}

// Load the state of the tree from cookies
function Tree_loadState() {
 var s = getCookie("yoffset")
 if(s != null) yoffset = s
 s = getCookie(this.id)
 if(s != null && s!="") this.setState(s)
}

// Save the state of the tree using cookies
function Tree_saveState() {
 setCookie(this.id, this.getState())
 setCookie("yoffset", yoffset)
}

// Get a string that contains the current state of each tree node
function Tree_getState() {
 var s = ""
 for(var i=0; i<this.nodes.length; ++i) 
  s += this.nodes[i].getState()
 return s
}

// Set the state of each tree node from a string
function Tree_setState(s) {
 for(var i=0; i<this.nodes.length; ++i)
  s = this.nodes[i].setState(s)
}

// Returns a new node ID
function Tree_getNodeID() {
 return this.nodeID++
}

// Returns the node corresponding to a specific ID
function Tree_getNode(id) {
 for(var i=0; i<this.nodes.length; ++i) {
  var node = this.nodes[i].findNode(id)
  if(node != null) return node
 }
 return null;
}

// Handles the clicking of a tree's marker
function Tree_stateChange(tree, nodeID) {
 // Get scroll position
 if(navigator.appName == "Netscape") 
  yoffset = window.pageYOffset
 else
  yoffset = window.document.body.scrollTop
 node = tree.getNode(nodeID)
 // Change the opened/closed state of the node
 if(node != null && node.nodes.length > 0) node.changeState()
}

// Sets the opened, closed, and leaf markers of the tree
function Tree_setMarkers(opened, closed, leaf) {
 this.openedMarker = opened
 this.closedMarker = closed
 this.leafMarker = leaf
}

// Returns the value of a cookie
function getCookie(name) {
 var cookie = removeBlanks(document.cookie)
 var nameValuePairs = cookie.split(";")
 for(var i=0; i<nameValuePairs.length; ++i) {
  var pairSplit = nameValuePairs[i].split("=")
  if(pairSplit[0] == name && pairSplit.length > 1) return pairSplit[1]
 }
 return null
}

// Sets the value of a cookie
function setCookie(name, value) {
 document.cookie = name+"="+value
}

// Removes any blanks in a string
function removeBlanks(s) {
 var temp = ""
 for(var i=0; i<s.length; ++i) {
  var c = s.charAt(i)
  if(c != " ") temp += c
 }
 return temp
}

// The Node object represents a node of a tree
// It is defined using the name (text) of the node, a hypertext link,
// a target window or frame, and a reference to the tree in which the 
// node is contained
function Node(name, link, target, tree) {
 this.name = name			// Text of the node
 this.link = link			// Hypertext link
 this.target = target			// Window or frame name
 this.tree = tree			// Reference to containing tree
 this.nodes = new Array()		// Array of sub-nodes
 this.opened = false			// Is the node expanded
 this.id = tree.getNodeID()		// A unique ID within the tree
 this.addNode = Node_addNode		// Add a sub-node
 this.display = Node_display		// Display the node
 this.open = Node_open			// Expand the node
 this.close = Node_close		// Close the node
 this.changeState = Node_changeState	// Change the expanded state of the node
 this.findNode = Node_findNode		// Returns a node with a specified ID
 this.getState = Node_getState		// Returns a string that describes the expanded state of the node & its sub-nodes
 this.setState = Node_setState		// Sets the expanded state of a node and its sub-nodes
}

// Add a sub-node
function Node_addNode(node) {
 this.nodes[this.nodes.length] = node
}

// Write the anchor tags associated with a node
function writeAnchor(treeref, id, marker, link, target, name) {
 document.write('<a href="javascript:Tree_stateChange('+treeref+","+id+')" class="treemarker">')
 document.write(marker)
 document.write('</a>')
 document.write('<a href="'+link+'"')
 if(target != "") document.write(' target="'+this.target+'"')
 document.write(' class="treenode">')
 document.write(name)
 document.write('</a>')
}

// Display the node
function Node_display() {
 if(this.nodes.length == 0) {
  document.write('<blockquote class="treenode">')
  writeAnchor(this.tree.variable, this.id, this.tree.leafMarker, this.link, this.target, this.name)
  document.writeln('</blockquote>')
 }else if(this.opened) {
  document.write('<blockquote class="treenode">')
  writeAnchor(this.tree.variable, this.id, this.tree.openedMarker, this.link, this.target, this.name)
  for(var i=0; i<this.nodes.length; ++i) this.nodes[i].display()
  document.writeln('</blockquote>')
 }else{
  document.write('<blockquote class="treenode">')
  writeAnchor(this.tree.variable, this.id, this.tree.closedMarker, this.link, this.target, this.name)
  document.writeln('</blockquote>')
 }
}

// Expand the node
function Node_open() {
 this.opened = true
 this.tree.saveState()
 var urls = window.location.href.split("#",1)
 window.location.reload()
}

// Close the node
function Node_close() {
 this.opened = false
 this.tree.saveState()
 window.location.reload()
}

// Change the expanded state of the node
function Node_changeState() {
 if(this.opened) this.close()
 else this.open()
}

// Returns a node with a specified ID
function Node_findNode(id) {
 if(this.id == id) return this
 for(var i=0; i<this.nodes.length; ++i) {
  var node = this.nodes[i].findNode(id)
  if(node != null) return node
 }
 return null
}

// Returns a string that describes the expanded state of the node & its sub-nodes
function Node_getState() {
 var s = ""
 if(this.opened) s += "1"
 else s += 0
 for(var i=0; i<this.nodes.length; ++i)
  s += this.nodes[i].getState() 
 return s
}

// Sets the expanded state of a node and its sub-nodes
function Node_setState(s) {
 if(s == null || s == "") return
 var bit = s.substring(0,1)
 if(bit == "1") this.opened = true
 else this.opened = false
 s = s.substring(1)
 for(var i=0; i<this.nodes.length; ++i)
  s = this.nodes[i].setState(s)
 return s
}
