"""
SharePoint MCP Server with Modern Azure AD Authentication
"""
import os
import logging
import asyncio
from functools import wraps
from typing import Optional
import base64
import mimetypes

from mcp.server import Server
from mcp.types import Resource, Tool, TextContent, ImageContent, EmbeddedResource
from pydantic import AnyUrl
import mcp.server.stdio

from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.client_context import ClientContext

from .auth import create_sharepoint_context

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize MCP server
app = Server("mcp-sharepoint")

# Global SharePoint context
ctx: Optional[ClientContext] = None


def ensure_context(func):
    """Decorator to ensure SharePoint context is available"""
    @wraps(func)
    async def wrapper(*args, **kwargs):
        global ctx
        if ctx is None:
            try:
                ctx = create_sharepoint_context()
                logger.info("SharePoint context initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize SharePoint context: {e}")
                raise RuntimeError(
                    f"SharePoint authentication failed: {e}. "
                    "Please check your environment variables and ensure:\n"
                    "1. SHP_TENANT_ID is set correctly\n"
                    "2. Your Azure AD app has the correct API permissions\n"
                    "3. If using a new tenant, make sure you're using modern auth (MSAL)"
                )
        return await func(*args, **kwargs)
    return wrapper


def get_document_library_path() -> str:
    """Get the document library path from environment"""
    return os.getenv("SHP_DOC_LIBRARY", "Shared Documents")


@app.list_resources()
async def list_resources() -> list[Resource]:
    """List available SharePoint resources"""
    return [
        Resource(
            uri=AnyUrl(f"sharepoint:///{get_document_library_path()}"),
            name=f"SharePoint Document Library: {get_document_library_path()}",
            mimeType="application/vnd.sharepoint.folder",
            description="Main SharePoint document library configured for this server"
        )
    ]


@app.list_tools()
async def list_tools() -> list[Tool]:
    """List available SharePoint tools"""
    return [
        Tool(
            name="List_SharePoint_Folders",
            description="List all folders in a specified directory or root of the document library",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Path to the folder (relative to document library root). Leave empty for root.",
                        "default": ""
                    }
                }
            }
        ),
        Tool(
            name="List_SharePoint_Documents",
            description="List all documents in a specified folder with metadata",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Path to the folder containing documents",
                        "default": ""
                    }
                },
                "required": []
            }
        ),
        Tool(
            name="Get_Document_Content",
            description="Get the content of a document (supports text extraction from PDF, Word, Excel, and text files)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the file (relative to document library root)"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="Upload_Document",
            description="Upload a new document to SharePoint",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Destination folder path"
                    },
                    "file_name": {
                        "type": "string",
                        "description": "Name of the file to create"
                    },
                    "content": {
                        "type": "string",
                        "description": "File content (text or base64 encoded for binary files)"
                    },
                    "is_binary": {
                        "type": "boolean",
                        "description": "Whether the content is base64 encoded binary",
                        "default": False
                    }
                },
                "required": ["folder_path", "file_name", "content"]
            }
        ),
        Tool(
            name="Update_Document",
            description="Update the content of an existing document",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the file to update"
                    },
                    "content": {
                        "type": "string",
                        "description": "New file content"
                    },
                    "is_binary": {
                        "type": "boolean",
                        "description": "Whether the content is base64 encoded binary",
                        "default": False
                    }
                },
                "required": ["file_path", "content"]
            }
        ),
        Tool(
            name="Delete_Document",
            description="Delete a document from SharePoint",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the file to delete"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="Create_Folder",
            description="Create a new folder in SharePoint",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Path where to create the folder"
                    },
                    "folder_name": {
                        "type": "string",
                        "description": "Name of the new folder"
                    }
                },
                "required": ["folder_path", "folder_name"]
            }
        ),
        Tool(
            name="Delete_Folder",
            description="Delete an empty folder from SharePoint",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Path to the folder to delete"
                    }
                },
                "required": ["folder_path"]
            }
        ),
        Tool(
            name="Get_SharePoint_Tree",
            description="Get a recursive tree view of SharePoint folder structure",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Starting folder path (leave empty for root)",
                        "default": ""
                    },
                    "max_depth": {
                        "type": "integer",
                        "description": "Maximum depth to traverse",
                        "default": 5
                    }
                },
                "required": []
            }
        ),
        Tool(
            name="Test_Connection",
            description="Test the SharePoint connection and authentication",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        )
    ]


@app.call_tool()
@ensure_context
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle tool execution"""
    
    try:
        if name == "Test_Connection":
            return await test_connection()
        elif name == "List_SharePoint_Folders":
            return await list_folders(arguments.get("folder_path", ""))
        elif name == "List_SharePoint_Documents":
            return await list_documents(arguments.get("folder_path", ""))
        elif name == "Get_Document_Content":
            return await get_document_content(arguments["file_path"])
        elif name == "Upload_Document":
            return await upload_document(
                arguments["folder_path"],
                arguments["file_name"],
                arguments["content"],
                arguments.get("is_binary", False)
            )
        elif name == "Update_Document":
            return await update_document(
                arguments["file_path"],
                arguments["content"],
                arguments.get("is_binary", False)
            )
        elif name == "Delete_Document":
            return await delete_document(arguments["file_path"])
        elif name == "Create_Folder":
            return await create_folder(arguments["folder_path"], arguments["folder_name"])
        elif name == "Delete_Folder":
            return await delete_folder(arguments["folder_path"])
        elif name == "Get_SharePoint_Tree":
            return await get_tree(
                arguments.get("folder_path", ""),
                arguments.get("max_depth", 5)
            )
        else:
            raise ValueError(f"Unknown tool: {name}")
            
    except Exception as e:
        logger.error(f"Tool '{name}' failed: {e}")
        return [TextContent(
            type="text",
            text=f"Error executing {name}: {str(e)}"
        )]


async def test_connection() -> list[TextContent]:
    """Test SharePoint connection"""
    try:
        web = ctx.web.get().execute_query()
        auth_method = os.getenv("SHP_AUTH_METHOD", "msal")
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully connected to SharePoint!\n\n"
                 f"Site Title: {web.title}\n"
                 f"Site URL: {web.url}\n"
                 f"Authentication Method: {auth_method.upper()}\n"
                 f"Tenant ID: {os.getenv('SHP_TENANT_ID')}\n\n"
                 f"Connection is working correctly with modern Azure AD authentication."
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"✗ Connection failed: {str(e)}\n\n"
                 f"This usually means:\n"
                 f"1. Your credentials are incorrect\n"
                 f"2. Your app doesn't have proper SharePoint permissions\n"
                 f"3. You're using legacy auth on a new tenant (set SHP_AUTH_METHOD=msal)"
        )]


async def list_folders(folder_path: str = "") -> list[TextContent]:
    """List folders in specified path"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}" if folder_path else doc_lib
        
        folder = ctx.web.get_folder_by_server_relative_path(full_path)
        folders = folder.folders.get().execute_query()
        
        folder_list = []
        for f in folders:
            folder_list.append(f"📁 {f.name}")
        
        result = f"Folders in '{full_path}':\n\n" + "\n".join(folder_list) if folder_list else f"No folders found in '{full_path}'"
        
        return [TextContent(type="text", text=result)]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error listing folders: {str(e)}")]


async def list_documents(folder_path: str = "") -> list[TextContent]:
    """List documents in specified folder"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}" if folder_path else doc_lib
        
        folder = ctx.web.get_folder_by_server_relative_path(full_path)
        files = folder.files.get().execute_query()
        
        file_list = []
        for f in files:
            size_kb = f.length / 1024
            file_list.append(f"📄 {f.name} ({size_kb:.2f} KB)")
        
        result = f"Documents in '{full_path}':\n\n" + "\n".join(file_list) if file_list else f"No documents found in '{full_path}'"
        
        return [TextContent(type="text", text=result)]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error listing documents: {str(e)}")]


async def get_document_content(file_path: str) -> list[TextContent]:
    """Get document content"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{file_path}"
        
        file = ctx.web.get_file_by_server_relative_path(full_path)
        content = file.read()
        
        # Determine if binary based on file extension
        ext = os.path.splitext(file_path)[1].lower()
        text_extensions = {'.txt', '.md', '.json', '.xml', '.html', '.csv', '.log'}
        
        if ext in text_extensions:
            # Text file
            text_content = content.decode('utf-8')
            return [TextContent(type="text", text=text_content)]
        else:
            # Binary file - return base64
            b64_content = base64.b64encode(content).decode('utf-8')
            return [TextContent(
                type="text",
                text=f"Binary file (base64 encoded):\n\n{b64_content[:200]}...\n\n"
                     f"Full content length: {len(b64_content)} characters"
            )]
            
    except Exception as e:
        return [TextContent(type="text", text=f"Error reading document: {str(e)}")]


async def upload_document(folder_path: str, file_name: str, content: str, is_binary: bool = False) -> list[TextContent]:
    """Upload a document"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}" if folder_path else doc_lib
        
        folder = ctx.web.get_folder_by_server_relative_path(full_path)
        
        if is_binary:
            file_content = base64.b64decode(content)
        else:
            file_content = content.encode('utf-8')
        
        uploaded_file = folder.upload_file(file_name, file_content).execute_query()
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully uploaded '{file_name}' to '{full_path}'"
        )]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error uploading document: {str(e)}")]


async def update_document(file_path: str, content: str, is_binary: bool = False) -> list[TextContent]:
    """Update a document"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{file_path}"
        
        if is_binary:
            file_content = base64.b64decode(content)
        else:
            file_content = content.encode('utf-8')
        
        file = ctx.web.get_file_by_server_relative_path(full_path)
        file.write(file_content).execute_query()
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully updated '{file_path}'"
        )]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error updating document: {str(e)}")]


async def delete_document(file_path: str) -> list[TextContent]:
    """Delete a document"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{file_path}"
        
        file = ctx.web.get_file_by_server_relative_path(full_path)
        file.delete_object().execute_query()
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully deleted '{file_path}'"
        )]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error deleting document: {str(e)}")]


async def create_folder(folder_path: str, folder_name: str) -> list[TextContent]:
    """Create a folder"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}" if folder_path else doc_lib
        
        parent_folder = ctx.web.get_folder_by_server_relative_path(full_path)
        new_folder = parent_folder.folders.add(folder_name).execute_query()
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully created folder '{folder_name}' in '{full_path}'"
        )]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error creating folder: {str(e)}")]


async def delete_folder(folder_path: str) -> list[TextContent]:
    """Delete a folder"""
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}"
        
        folder = ctx.web.get_folder_by_server_relative_path(full_path)
        folder.delete_object().execute_query()
        
        return [TextContent(
            type="text",
            text=f"✓ Successfully deleted folder '{folder_path}'"
        )]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error deleting folder: {str(e)}")]


async def get_tree(folder_path: str = "", max_depth: int = 5, current_depth: int = 0) -> list[TextContent]:
    """Get folder tree structure"""
    if current_depth >= max_depth:
        return [TextContent(type="text", text="Max depth reached")]
    
    try:
        doc_lib = get_document_library_path()
        full_path = f"{doc_lib}/{folder_path}" if folder_path else doc_lib
        
        folder = ctx.web.get_folder_by_server_relative_path(full_path)
        folders = folder.folders.get().execute_query()
        
        indent = "  " * current_depth
        tree_lines = [f"{indent}📁 {folder_path or 'Root'}"]
        
        for f in folders:
            sub_path = f"{folder_path}/{f.name}" if folder_path else f.name
            sub_tree = await get_tree(sub_path, max_depth, current_depth + 1)
            tree_lines.append(sub_tree[0].text)
        
        return [TextContent(type="text", text="\n".join(tree_lines))]
        
    except Exception as e:
        return [TextContent(type="text", text=f"Error getting tree: {str(e)}")]


async def main():
    """Main entry point"""
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
