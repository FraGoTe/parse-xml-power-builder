//wf_parse_node(node, treeview)
// Arguments
oleobject aole_node // The node to process
long al_treeview_parent // The treeview item to add
nodes to
// local variables
string ls_node_type // The type of the current node
string ls_node_name // The name of the current node
string ls_node_value // The value of the current node
oleobject lole_attribute_node_list // The attributes for
this node
oleobject lole_attribute_node // An attribute for this node
oleobject lole_child_node_list // The child nodes for this
node
oleobject lole_child_node // The current child node
long ll_max_attribute_nodes // The number of attribute
nodes
long ll_max_child_nodes // The number of child nodes
long ll_attribute_idx // A counter
long ll_child_idx // A counter
string ls_label // The label for the treeview
long ll_treeview // The newly inserted treeview item
// Determine the node's name, type and value
ls_node_name = aole_node.nodename
ls_node_type = Upper(aole_node.nodetypestring)
ls_node_value = String(aole_node.nodevalue)
IF IsNull(ls_node_value) THEN
ls_node_value = "<NULL>"
END IF
// Add this node to the treeview
// I am using a picture index of 1 however
// you could use the nodetypestring to determine
// the node‚s type and use an appropriate picture
ls_label = ls_node_name + " (" + ls_node_value + ")"
ll_treeview = tv_xml.InsertItemLast (al_treeview_parent,
ls_label, )
// If this node has attributes process them
lole_attribute_node_list = aole_node.attributes
IF IsValid(lole_attribute_node_list) THEN
ll_max_attribute_nodes = lole_attribute_node_list.length
ELSE
ll_max_attribute_nodes = 0
END IF
FOR ll_attribute_idx = 0 TO ll_max_attribute_nodes - 1
lole_attribute_node =
lole_attribute_node_list.Item(ll_attribute_idx)
wf_parse_node (ll_treeview, lole_attribute_node) /* RECURSIVE
*/
NEXT
// If this node is an element and it has children process them
IF ls_node_type = "ELEMENT" THEN
lole_child_node_list = aole_node.childNodes
IF IsValid(lole_child_node_list) THEN
ll_max_child_nodes = lole_child_node_list.length
ELSE
ll_max_child_nodes = 0
END IF
FOR ll_child_idx = 0 TO ll_max_child_nodes - 1
lole_child_node =
lole_child_node_list.Item(ll_child_idx)
wf_parse_node(ll_treeview, lole_child_node) /* RECURSIVE
*/
NEXT
END IF
