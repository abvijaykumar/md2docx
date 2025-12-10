# -*- coding: utf-8 -*-
# MIT License
#
# Copyright (c) 2025 A B Vijay Kumar
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import re
import os
import sys
import glob
import argparse
import xml.etree.ElementTree as ET
from xml.dom import minidom
import base64
import zlib
import urllib.parse
import math

class MermaidToDrawioConverter:
    def __init__(self):
        self.node_counter = 1
        self.edge_counter = 1
        
    def generate_node_id(self):
        """Generate unique node ID"""
        node_id = "node" + str(self.node_counter)
        self.node_counter += 1
        return node_id
    
    def generate_edge_id(self):
        """Generate unique edge ID"""
        edge_id = "edge" + str(self.edge_counter)
        self.edge_counter += 1
        return edge_id
    
    def reset_counters(self):
        """Reset ID counters for new diagram"""
        self.node_counter = 1
        self.edge_counter = 1
    
    def parse_node_shape(self, node_text):
        """Parse node shape and return appropriate Draw.io style"""
        base_style = "whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;"
        
        # Handle different node shapes with proper precedence (longer patterns first)
        if node_text.startswith('[[') and node_text.endswith(']]'):
            # Subroutine: A[[text]]
            return node_text[2:-2], f"rounded=1;strokeWidth=2;{base_style}"
        elif node_text.startswith('{{') and node_text.endswith('}}'):
            # Hexagon: A{{text}}
            return node_text[2:-2], f"shape=hexagon;perimeter=hexagonPerimeter2;fillColor=#fff2cc;strokeColor=#d6b656;{base_style}"
        elif node_text.startswith('((') and node_text.endswith('))'):
            # Circle: A((text))
            return node_text[2:-2], f"ellipse;aspect=fixed;fillColor=#f8cecc;strokeColor=#b85450;{base_style}"
        elif node_text.startswith('[(') and node_text.endswith(')]'):
            # Database: A[(text)]
            return node_text[2:-2], f"shape=cylinder3;fillColor=#e1d5e7;strokeColor=#9673a6;{base_style}"
        elif node_text.startswith('>') and node_text.endswith(']'):
            # Flag: A>text]
            return node_text[1:-1], f"shape=parallelogram;perimeter=parallelogramPerimeter;fillColor=#d5e8d4;strokeColor=#82b366;{base_style}"
        elif node_text.startswith('[') and node_text.endswith(']'):
            # Rectangle: A[text] (must come after other bracket types)
            return node_text[1:-1], f"rounded=1;{base_style}"
        elif node_text.startswith('(') and node_text.endswith(')'):
            # Round node: A(text) (must come after double parentheses)
            return node_text[1:-1], f"ellipse;fillColor=#ffe6cc;strokeColor=#d79b00;{base_style}"
        elif node_text.startswith('{') and node_text.endswith('}'):
            # Diamond/Decision: A{text} (must come after double braces)
            return node_text[1:-1], f"rhombus;fillColor=#fff2cc;strokeColor=#d6b656;{base_style}"
        else:
            # Default rectangle
            return node_text, f"rounded=1;{base_style}"
    
    def parse_arrow_type(self, arrow_text):
        """Parse arrow type and return Draw.io style"""
        style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#6c8ebf;"
        
        # Dotted/Dashed arrows
        if '.-' in arrow_text or '..' in arrow_text:
            style += "dashed=1;dashPattern=5 5;"
        
        # Thick arrows
        if '==' in arrow_text:
            style += "strokeWidth=3;"
        else:
            style += "strokeWidth=2;"
        
        # Arrow head types
        if arrow_text.endswith('->'):
            style += "endArrow=classic;endFill=1;"
        elif arrow_text.endswith('-'):
            style += "endArrow=none;"
        elif arrow_text.endswith('o'):
            style += "endArrow=oval;endFill=0;"
        elif arrow_text.endswith('x'):
            style += "endArrow=cross;endFill=1;"
        else:
            style += "endArrow=classic;endFill=1;"
        
        return style
    
    def parse_mermaid_flowchart(self, mermaid_text):
        """Parse mermaid flowchart and extract nodes and edges"""
        nodes = {}
        edges = []
        
        lines = mermaid_text.strip().split('\n')
        
        # Detect flow direction
        self.flow_direction = 'TD'  # Default
        for line in lines:
            line = line.strip()
            if line.startswith('graph') or line.startswith('flowchart'):
                if 'LR' in line:
                    self.flow_direction = 'LR'
                elif 'RL' in line:
                    self.flow_direction = 'RL'
                elif 'BT' in line:
                    self.flow_direction = 'BT'
                else:
                    self.flow_direction = 'TD'
                continue
            if not line:
                continue
            
            # Handle mermaid format: A -->|label| B by converting to A --> B |label|
            edge_label_match = re.search(r'([-.=]{1,3}>?[ox]?|[-.=]{1,3})\|([^|]+)\|', line)
            extracted_label = ''
            if edge_label_match:
                arrow_part = edge_label_match.group(1)
                extracted_label = edge_label_match.group(2)
                # Replace -->|label| with --> and store the label
                line = re.sub(r'([-.=]{1,3}>?[ox]?|[-.=]{1,3})\|[^|]+\|', arrow_part, line)
            
            # Enhanced pattern to handle all mermaid arrow types and node shapes
            # Order matters: check complex shapes first, then simple ones
            arrow_pattern = r'([A-Za-z0-9_]+)(?:(\[\[[^\]]+\]\]|\{\{[^\}]+\}\}|\(\([^\)]+\)\)|\[\([^\]]+\)\]|>[^\]]+\]|\[[^\]]+\]|\([^\)]+\)|\{[^\}]+\}))?\s*([-.=]{1,3}>?[ox]?|[-.=]{1,3})\s*([A-Za-z0-9_]+)(?:(\[\[[^\]]+\]\]|\{\{[^\}]+\}\}|\(\([^\)]+\)\)|\[\([^\)]+\)\]|>[^\]]+\]|\[[^\]]+\]|\([^\)]+\)|\{[^\}]+\}))?(?:\s*\|\s*([^|]+)\s*\|)?'
            
            match = re.search(arrow_pattern, line)
            
            if match:
                from_node = match.group(1)
                from_shape = match.group(2)
                arrow_type = match.group(3)
                to_node = match.group(4)
                to_shape = match.group(5)
                edge_label = extracted_label if extracted_label else (match.group(6) if match.group(6) else '')
                
                # Parse from node
                if from_node not in nodes:
                    if from_shape:
                        label, style = self.parse_node_shape(from_shape)
                    else:
                        label, style = from_node, "rounded=1;whiteSpace=wrap;html=1;"
                    nodes[from_node] = {
                        'id': self.generate_node_id(), 
                        'label': label,
                        'style': style
                    }
                elif from_shape and nodes[from_node]['label'] == from_node:
                    # Update with shape info if available
                    label, style = self.parse_node_shape(from_shape)
                    nodes[from_node]['label'] = label
                    nodes[from_node]['style'] = style
                
                # Parse to node
                if to_node not in nodes:
                    if to_shape:
                        label, style = self.parse_node_shape(to_shape)
                    else:
                        label, style = to_node, "rounded=1;whiteSpace=wrap;html=1;"
                    nodes[to_node] = {
                        'id': self.generate_node_id(), 
                        'label': label,
                        'style': style
                    }
                elif to_shape and nodes[to_node]['label'] == to_node:
                    # Update with shape info if available
                    label, style = self.parse_node_shape(to_shape)
                    nodes[to_node]['label'] = label
                    nodes[to_node]['style'] = style
                
                # Parse arrow style
                arrow_style = self.parse_arrow_type(arrow_type)
                
                # Add edge
                edge = {
                    'id': self.generate_edge_id(),
                    'from': nodes[from_node]['id'],
                    'to': nodes[to_node]['id'],
                    'label': edge_label.strip() if edge_label else '',
                    'style': arrow_style
                }
                edges.append(edge)
            
            # Handle standalone node definitions with shapes
            else:
                # Order matters: check complex shapes first
                standalone_pattern = r'([A-Za-z0-9_]+)(\[\[[^\]]+\]\]|\{\{[^\}]+\}\}|\(\([^\)]+\)\)|\[\([^\]]+\)\]|>[^\]]+\]|\[[^\]]+\]|\([^\)]+\)|\{[^\}]+\})'
                match = re.search(standalone_pattern, line)
                if match:
                    node_name = match.group(1)
                    node_shape = match.group(2)
                    
                    if node_name not in nodes:
                        label, style = self.parse_node_shape(node_shape)
                        nodes[node_name] = {
                            'id': self.generate_node_id(), 
                            'label': label,
                            'style': style
                        }
        
        return nodes, edges
    
    def parse_sequence_arrow_style(self, arrow_text):
        """Parse sequence diagram arrow and return Draw.io style"""
        base_style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;"
        
        # Synchronous messages: ->
        if arrow_text == '->':
            return base_style + "endArrow=classic;"
        
        # Asynchronous messages: ->>
        elif arrow_text == '->>':
            return base_style + "endArrow=open;dashed=1;"
        
        # Dotted messages: -.->
        elif '.->' in arrow_text:
            return base_style + "endArrow=classic;dashed=1;"
        
        # Activation/Deactivation: +, -
        elif arrow_text == '+':
            return base_style + "endArrow=classic;strokeWidth=2;"
        elif arrow_text == '-':
            return base_style + "endArrow=classic;strokeWidth=2;"
        
        # Return messages: -->>
        elif '-->' in arrow_text:
            return base_style + "endArrow=classic;dashed=1;"
        
        # Cross/X ending: -x
        elif arrow_text.endswith('x'):
            return base_style + "endArrow=cross;"
        
        # Default
        else:
            return base_style + "endArrow=classic;"
    
    def parse_mermaid_sequence(self, mermaid_text):
        """Parse mermaid sequence diagram"""
        participants = {}
        messages = []
        
        lines = mermaid_text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('sequenceDiagram'):
                continue
            
            # Parse participant definitions
            if line.startswith('participant') or line.startswith('actor'):
                parts = line.split(' as ', 1)
                if len(parts) == 2:
                    name = parts[0].replace('participant', '').replace('actor', '').strip()
                    label = parts[1].strip()
                else:
                    name = line.replace('participant', '').replace('actor', '').strip()
                    label = name
                
                if name not in participants:
                    participants[name] = {'id': self.generate_node_id(), 'label': label}
            
            # Parse messages with comprehensive arrow support
            elif any(arrow in line for arrow in ['->', '->>', '-.->',  '-->', '-x', '+', '-']):
                # Enhanced pattern for all sequence diagram arrows
                msg_pattern = r'([A-Za-z0-9_]+)\s*(->>?|\.->|-->|->|-x|\+|-)\s*([A-Za-z0-9_]+)\s*:\s*(.+)'
                match = re.search(msg_pattern, line)
                
                if match:
                    from_part = match.group(1)
                    arrow = match.group(2)
                    to_part = match.group(3)
                    message = match.group(4)
                    
                    # Add participants if not defined
                    if from_part not in participants:
                        participants[from_part] = {'id': self.generate_node_id(), 'label': from_part}
                    if to_part not in participants:
                        participants[to_part] = {'id': self.generate_node_id(), 'label': to_part}
                    
                    # Get appropriate Draw.io style for arrow type
                    arrow_style = self.parse_sequence_arrow_style(arrow)
                    
                    messages.append({
                        'id': self.generate_edge_id(),
                        'from': participants[from_part]['id'],
                        'to': participants[to_part]['id'],
                        'label': message.strip(),
                        'style': arrow_style
                    })
            
            # Handle notes, loops, and other sequence elements
            elif line.startswith('Note'):
                # Notes could be added as text boxes in the future
                pass
            elif line.startswith('loop') or line.startswith('alt') or line.startswith('opt'):
                # Complex flow control - could be handled in advanced version
                pass
        
        return participants, messages
    
    def parse_mermaid_er(self, mermaid_text):
        """Parse mermaid ER diagram"""
        entities = {}
        relationships = []
        
        lines = mermaid_text.strip().split('\n')
        current_entity = None
        in_entity_block = False
        
        for line in lines:
            original_line = line
            line = line.strip()
            if not line or line.startswith('erDiagram'):
                continue
            
            # Check if we're entering an entity block
            if line.endswith('{'):
                current_entity = line.rstrip(' {').strip()
                in_entity_block = True
                if current_entity not in entities:
                    entities[current_entity] = {
                        'id': self.generate_node_id(),
                        'label': current_entity,
                        'attributes': []
                    }
                continue
            
            # Check if we're exiting an entity block
            if line == '}':
                in_entity_block = False
                current_entity = None
                continue
            
            # Parse attributes within entity block
            if in_entity_block and current_entity and current_entity in entities:
                attr = line.strip()
                if attr:
                    entities[current_entity]['attributes'].append(attr)
                continue
            
            # Parse entity definitions without blocks (just entity names)
            if not in_entity_block and line and not ('||--' in line or '}o--' in line or '||..o{' in line):
                # This might be an entity name
                entity_name = line.strip()
                if entity_name and not any(char in entity_name for char in [':', '||', '--', '}o', '{']):
                    if entity_name not in entities:
                        entities[entity_name] = {
                            'id': self.generate_node_id(),
                            'label': entity_name,
                            'attributes': []
                        }
                    current_entity = entity_name
                continue
            
            # Parse relationships with comprehensive ER notation support
            if any(rel in line for rel in ['||--', '}o--', '||..', '}|--', '|o--', 'o|--']):
                # Enhanced pattern for all ER relationship types
                rel_pattern = r'([A-Za-z0-9_]+)\s*([\|\}][|\}o][-\.]{2,3}[o\|\}][\{\|])\s*([A-Za-z0-9_]+)\s*:\s*(.+)'
                match = re.search(rel_pattern, line)
                
                if match:
                    from_entity = match.group(1)
                    relationship = match.group(2)
                    to_entity = match.group(3)
                    label = match.group(4)
                    
                    # Add entities if not defined
                    if from_entity not in entities:
                        entities[from_entity] = {
                            'id': self.generate_node_id(),
                            'label': from_entity,
                            'attributes': []
                        }
                    if to_entity not in entities:
                        entities[to_entity] = {
                            'id': self.generate_node_id(),
                            'label': to_entity,
                            'attributes': []
                        }
                    
                    # Determine relationship style based on notation
                    rel_style = self.parse_er_relationship_style(relationship)
                    
                    relationships.append({
                        'id': self.generate_edge_id(),
                        'from': entities[from_entity]['id'],
                        'to': entities[to_entity]['id'],
                        'label': label,
                        'style': rel_style
                    })
        
        return entities, relationships
    
    def parse_er_relationship_style(self, relationship_text):
        """Parse ER relationship notation and return Draw.io style"""
        base_style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;"
        
        # Dotted relationships (identifying/non-identifying)
        if '..' in relationship_text:
            base_style += "dashed=1;"
        
        # Determine cardinality markers
        if relationship_text.startswith('||'):
            # One-to relationship
            base_style += "startArrow=ERone;"
        elif relationship_text.startswith('}|'):
            # Many-to relationship
            base_style += "startArrow=ERmany;"
        elif relationship_text.startswith('|o'):
            # Zero-or-one relationship
            base_style += "startArrow=ERzeroToOne;"
        elif relationship_text.startswith('}o'):
            # Zero-or-many relationship
            base_style += "startArrow=ERzeroToMany;"
        
        if relationship_text.endswith('||'):
            # To-one relationship
            base_style += "endArrow=ERone;"
        elif relationship_text.endswith('|{'):
            # To-many relationship
            base_style += "endArrow=ERmany;"
        elif relationship_text.endswith('o|'):
            # To-zero-or-one relationship
            base_style += "endArrow=ERzeroToOne;"
        elif relationship_text.endswith('o{'):
            # To-zero-or-many relationship
            base_style += "endArrow=ERzeroToMany;"
        
        return base_style
    
    def parse_mermaid_state(self, mermaid_text):
        """Parse mermaid state diagram"""
        states = {}
        transitions = []
        
        lines = mermaid_text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('stateDiagram'):
                continue
            
            # Parse state transitions: [*] --> State or State1 --> State2: Label
            transition_pattern = r'(\[?\*?\]?[A-Za-z0-9_]*)\s*-->\s*(\[?\*?\]?[A-Za-z0-9_]*)\s*(?::\s*(.+))?'
            match = re.search(transition_pattern, line)
            
            if match:
                from_state = match.group(1).strip()
                to_state = match.group(2).strip()
                label = match.group(3).strip() if match.group(3) else ''
                
                # Handle initial/final states [*]
                if from_state == '[*]':
                    from_state = 'Start'
                elif from_state == '*':
                    from_state = 'Start'
                    
                if to_state == '[*]':
                    to_state = 'End'
                elif to_state == '*':
                    to_state = 'End'
                
                # Add states if not exists
                if from_state not in states:
                    display_name = '●' if from_state == 'Start' else ('◉' if from_state == 'End' else from_state)
                    states[from_state] = {'id': self.generate_node_id(), 'label': display_name}
                
                if to_state not in states:
                    display_name = '●' if to_state == 'Start' else ('◉' if to_state == 'End' else to_state)
                    states[to_state] = {'id': self.generate_node_id(), 'label': display_name}
                
                # Add transition
                transition = {
                    'id': self.generate_edge_id(),
                    'from': states[from_state]['id'],
                    'to': states[to_state]['id'],
                    'label': label,
                    'style': 'solid'
                }
                transitions.append(transition)
        
        return states, transitions
    
    def create_drawio_node(self, node_id, label, x, y, width=120, height=60, style="rounded=1;whiteSpace=wrap;html=1;"):
        """Create a Draw.io XML node"""
        cell = ET.Element('mxCell')
        cell.set('id', node_id)
        cell.set('value', label)
        cell.set('style', style)
        cell.set('vertex', '1')
        cell.set('parent', '1')
        
        geometry = ET.SubElement(cell, 'mxGeometry')
        geometry.set('x', str(x))
        geometry.set('y', str(y))
        geometry.set('width', str(width))
        geometry.set('height', str(height))
        geometry.set('as', 'geometry')
        
        return cell
    
    def create_drawio_edge(self, edge_id, source_id, target_id, label="", style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;"):
        """Create a Draw.io XML edge"""
        cell = ET.Element('mxCell')
        cell.set('id', edge_id)
        cell.set('value', label)
        cell.set('style', style)
        cell.set('edge', '1')
        cell.set('parent', '1')
        cell.set('source', source_id)
        cell.set('target', target_id)
        
        geometry = ET.SubElement(cell, 'mxGeometry')
        geometry.set('relative', '1')
        geometry.set('as', 'geometry')
        
        return cell
    
    def calculate_positions(self, nodes, edges=None, diagram_type='flowchart'):
        """Calculate positions for nodes based on diagram type and connections"""
        positions = {}
        node_list = list(nodes.keys())
        
        if diagram_type == 'sequence':
            # Arrange participants horizontally
            x_start = 100
            y_pos = 80
            x_spacing = 250
            
            for i, node in enumerate(node_list):
                positions[nodes[node]['id']] = {
                    'x': x_start + (i * x_spacing),
                    'y': y_pos,
                    'width': 180,
                    'height': 80
                }
        
        elif diagram_type == 'er':
            # Arrange entities in a grid with better spacing
            cols = max(2, min(4, math.ceil(math.sqrt(len(node_list)))))
            x_spacing = 300
            y_spacing = 250
            
            for i, node in enumerate(node_list):
                row = i // cols
                col = i % cols
                positions[nodes[node]['id']] = {
                    'x': 100 + (col * x_spacing),
                    'y': 100 + (row * y_spacing),
                    'width': 220,
                    'height': max(120, 50 + len(nodes[node].get('attributes', [])) * 25)
                }
        
        elif diagram_type == 'state':
            # Arrange states in a flow-like layout with better spacing
            cols = min(4, max(2, len(node_list)))
            x_spacing = 250
            y_spacing = 180
            
            for i, node in enumerate(node_list):
                row = i // cols
                col = i % cols
                positions[nodes[node]['id']] = {
                    'x': 100 + (col * x_spacing),
                    'y': 100 + (row * y_spacing),
                    'width': 180,
                    'height': 90
                }
        
        else:  # flowchart - improved layout
            # Try to create a more logical flow layout
            if edges and len(edges) > 0:
                # Use a simple hierarchical layout based on connections
                levels = self._calculate_node_levels(nodes, edges)
                positions = self._layout_by_levels(nodes, levels)
            else:
                # Fallback to grid layout with better spacing
                cols = max(2, min(5, math.ceil(math.sqrt(len(node_list)))))
                x_spacing = 280
                y_spacing = 180
                
                for i, node in enumerate(node_list):
                    row = i // cols
                    col = i % cols
                    positions[nodes[node]['id']] = {
                        'x': 100 + (col * x_spacing),
                        'y': 100 + (row * y_spacing),
                        'width': 200,
                        'height': 100
                    }
        
        return positions
    
    def _calculate_node_levels(self, nodes, edges):
        """Calculate hierarchical levels using topological sort"""
        # Build adjacency lists
        graph = {node_data['id']: [] for node_data in nodes.values()}
        in_degree = {node_data['id']: 0 for node_data in nodes.values()}
        
        for edge in edges:
            graph[edge['from']].append(edge['to'])
            in_degree[edge['to']] += 1
        
        # Topological sort with level assignment
        levels = {}
        queue = [node_id for node_id, degree in in_degree.items() if degree == 0]
        current_level = 0
        
        while queue:
            next_queue = []
            
            # Process all nodes at current level
            for node_id in queue:
                levels[node_id] = current_level
                
                # Reduce in-degree of neighbors
                for neighbor in graph[node_id]:
                    in_degree[neighbor] -= 1
                    if in_degree[neighbor] == 0:
                        next_queue.append(neighbor)
            
            queue = next_queue
            current_level += 1
        
        # Handle remaining nodes (cycles or disconnected)
        for node_id in in_degree:
            if node_id not in levels:
                levels[node_id] = current_level
        
        return levels
    
    def _layout_by_levels(self, nodes, levels):
        """Layout nodes by hierarchical levels with proper flow direction"""
        positions = {}
        
        # Group nodes by level
        level_groups = {}
        for node_id, level in levels.items():
            if level not in level_groups:
                level_groups[level] = []
            level_groups[level].append(node_id)
        
        # Detect flow direction from mermaid syntax
        flow_direction = self._detect_flow_direction()
        
        if flow_direction in ['TD', 'TB']:  # Top to Bottom
            y_spacing = 180
            x_spacing = 220
            start_x = 100
            start_y = 80
            
            for level, node_ids in sorted(level_groups.items()):
                y_pos = start_y + (level * y_spacing)
                
                # Center nodes horizontally
                total_width = (len(node_ids) - 1) * x_spacing
                start_x_level = start_x + max(0, (800 - total_width) // 2)
                
                for i, node_id in enumerate(node_ids):
                    positions[node_id] = {
                        'x': start_x_level + (i * x_spacing),
                        'y': y_pos,
                        'width': 180,
                        'height': 80
                    }
        
        else:  # Left to Right (LR)
            x_spacing = 220
            y_spacing = 150
            start_x = 80
            start_y = 100
            
            for level, node_ids in sorted(level_groups.items()):
                x_pos = start_x + (level * x_spacing)
                
                # Center nodes vertically
                total_height = (len(node_ids) - 1) * y_spacing
                start_y_level = start_y + max(0, (600 - total_height) // 2)
                
                for i, node_id in enumerate(node_ids):
                    positions[node_id] = {
                        'x': x_pos,
                        'y': start_y_level + (i * y_spacing),
                        'width': 180,
                        'height': 80
                    }
        
        return positions
    
    def _detect_flow_direction(self):
        """Detect flow direction from mermaid diagram"""
        # This would be set during parsing, defaulting to TD
        return getattr(self, 'flow_direction', 'TD')
    
    def convert_mermaid_to_drawio(self, mermaid_text, diagram_name="Diagram"):
        """Convert mermaid diagram to Draw.io XML"""
        self.reset_counters()
        
        # Detect diagram type
        if 'sequenceDiagram' in mermaid_text:
            diagram_type = 'sequence'
            nodes, edges = self.parse_mermaid_sequence(mermaid_text)
        elif 'erDiagram' in mermaid_text:
            diagram_type = 'er'
            nodes, edges = self.parse_mermaid_er(mermaid_text)
        elif 'stateDiagram' in mermaid_text:
            diagram_type = 'state'
            nodes, edges = self.parse_mermaid_state(mermaid_text)
        else:
            diagram_type = 'flowchart'
            nodes, edges = self.parse_mermaid_flowchart(mermaid_text)
        
        # Calculate positions
        positions = self.calculate_positions(nodes, edges, diagram_type)
        
        # Create Draw.io XML structure
        root = ET.Element('mxfile', host="app.diagrams.net", modified="2024-01-01T00:00:00.000Z", agent="mmd2drawio", etag="generated", version="22.1.0")
        diagram = ET.SubElement(root, 'diagram', id="diagram1", name=diagram_name)
        
        # Create mxGraphModel
        graph_model = ET.SubElement(diagram, 'mxGraphModel')
        graph_model.set('dx', '1422')
        graph_model.set('dy', '794')
        graph_model.set('grid', '1')
        graph_model.set('gridSize', '10')
        graph_model.set('guides', '1')
        graph_model.set('tooltips', '1')
        graph_model.set('connect', '1')
        graph_model.set('arrows', '1')
        graph_model.set('fold', '1')
        graph_model.set('page', '1')
        graph_model.set('pageScale', '1')
        graph_model.set('pageWidth', '827')
        graph_model.set('pageHeight', '1169')
        graph_model.set('math', '0')
        graph_model.set('shadow', '0')
        
        # Create root cell
        mx_cell_root = ET.SubElement(graph_model, 'root')
        
        # Add default cells
        cell0 = ET.SubElement(mx_cell_root, 'mxCell', id='0')
        cell1 = ET.SubElement(mx_cell_root, 'mxCell', id='1', parent='0')
        
        # Add nodes
        for node_name, node_data in nodes.items():
            node_id = node_data['id']
            pos = positions[node_id]
            
            if diagram_type == 'er':
                # Create entity with attributes
                label = node_data['label']
                if node_data.get('attributes'):
                    label += '\\n' + '\\n'.join(node_data['attributes'])
                
                style = "swimlane;fontStyle=0;childLayout=stackLayout;horizontal=1;startSize=30;horizontalStack=0;resizeParent=1;resizeParentMax=0;resizeLast=0;collapsible=1;marginBottom=0;"
                cell = self.create_drawio_node(
                    node_id, label, pos['x'], pos['y'], 
                    pos['width'], pos['height'], style
                )
            else:
                # Use custom style if available
                node_style = node_data.get('style', 'rounded=1;whiteSpace=wrap;html=1;')
                cell = self.create_drawio_node(
                    node_id, node_data['label'], pos['x'], pos['y'],
                    pos['width'], pos['height'], node_style
                )
            
            mx_cell_root.append(cell)
        
        # Add edges
        for edge in edges:
            # Use custom style if available, otherwise use default
            edge_style = edge.get('style', "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;")
            
            cell = self.create_drawio_edge(
                edge['id'], edge['from'], edge['to'], edge.get('label', ''), edge_style
            )
            mx_cell_root.append(cell)
        
        return root
    
    def save_drawio_file(self, xml_root, output_file):
        """Save XML to Draw.io file"""
        # Create pretty XML
        rough_string = ET.tostring(xml_root, 'unicode')
        reparsed = minidom.parseString(rough_string)
        pretty_xml = reparsed.toprettyxml(indent="  ")
        
        # Remove extra newlines
        pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
    
    def convert_single_file(self, mmd_file, output_file):
        """Convert single mermaid file to Draw.io"""
        with open(mmd_file, 'r', encoding='utf-8') as f:
            mermaid_content = f.read()
        
        diagram_name = os.path.splitext(os.path.basename(mmd_file))[0]
        xml_root = self.convert_mermaid_to_drawio(mermaid_content, diagram_name)
        self.save_drawio_file(xml_root, output_file)
    
    def convert_multiple_files(self, mmd_files, output_file):
        """Convert multiple mermaid files to single Draw.io file with multiple pages"""
        # Create root mxfile element
        root = ET.Element('mxfile', host="app.diagrams.net", modified="2024-01-01T00:00:00.000Z", agent="mmd2drawio", etag="generated", version="22.1.0")
        
        for i, mmd_file in enumerate(mmd_files):
            with open(mmd_file, 'r', encoding='utf-8') as f:
                mermaid_content = f.read()
            
            diagram_name = os.path.splitext(os.path.basename(mmd_file))[0]
            
            # Convert to Draw.io format
            single_diagram_root = self.convert_mermaid_to_drawio(mermaid_content, diagram_name)
            
            # Extract the diagram element and add to root
            diagram = single_diagram_root.find('diagram')
            diagram.set('id', "diagram" + str(i + 1))
            root.append(diagram)
        
        self.save_drawio_file(root, output_file)
    
    def process_mmd_files(self, input_path, output_dir=None, combine=False):
        """Process mermaid files in a directory"""
        if os.path.isfile(input_path):
            # Single file
            mmd_files = [input_path]
        else:
            # Directory
            mmd_files = sorted(glob.glob(os.path.join(input_path, '*.mmd')))
        
        if not mmd_files:
            print("No .mmd files found in " + str(input_path))
            return
        
        # Set up output directory
        if output_dir is None:
            if os.path.isfile(input_path):
                output_dir = os.path.dirname(input_path)
                if not output_dir:
                    output_dir = '.'
            else:
                output_dir = input_path
        
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        if combine:
            # Create single Draw.io file with multiple pages
            output_file = os.path.join(output_dir, "combined.drawio")
            self.convert_multiple_files(mmd_files, output_file)
            print("Created combined Draw.io file: " + output_file)
        else:
            # Create separate Draw.io file for each mermaid file
            for mmd_file in mmd_files:
                filename = os.path.splitext(os.path.basename(mmd_file))[0]
                output_file = os.path.join(output_dir, filename + ".drawio")
                self.convert_single_file(mmd_file, output_file)
                print("Converted: " + output_file)

def main():
    parser = argparse.ArgumentParser(
        description='Convert mermaid diagrams (.mmd files) to Draw.io format',
        epilog='''
Examples:
  python mmd2drawio.py ./diagrams/                    # Convert all .mmd files to separate .drawio files
  python mmd2drawio.py -c ./diagrams/                 # Combine all .mmd files into one .drawio file
  python mmd2drawio.py -o ./output/ ./diagrams/       # Specify output directory
  python mmd2drawio.py diagram.mmd                    # Convert single file
        '''
    )
    
    parser.add_argument('input_path', help='Path to .mmd file or directory containing .mmd files')
    parser.add_argument('-c', '--combine', action='store_true', 
                        help='Combine all mermaid files into a single Draw.io file with multiple pages')
    parser.add_argument('-o', '--output-dir', 
                        help='Output directory for Draw.io files (default: same as input directory)')
    
    args = parser.parse_args()
    
    converter = MermaidToDrawioConverter()
    converter.process_mmd_files(args.input_path, args.output_dir, args.combine)
    print("Conversion completed!")

if __name__ == "__main__":
    main()