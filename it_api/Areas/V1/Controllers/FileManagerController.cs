using Syncfusion.EJ2.FileManager.PhysicalFileProvider;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using Syncfusion.EJ2.FileManager.Base2;
using Vue.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace it_template.Areas.V1.Controllers
{

    [Area("V1")]
    public class FileManagerController : Controller
    {
        public PhysicalFileProvider operation;
        public string basePath;
        string root = "upload";
        protected readonly ItContext _context;
        public FileManagerController(ItContext context, IConfiguration configuration) : base()
        {
            _context = context;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
            basePath = configuration["Source:Path_Private"];
            operation = new PhysicalFileProvider();
            operation.RootFolder(basePath + "\\" + root);
        }
        public object FileOperations([FromBody] FileManagerDirectoryContent args)
        {
            // this.operation.SetRules(GetRules());
            if (args.Action == "delete" || args.Action == "rename")
            {
                if (args.TargetPath == null && args.Path == "")
                {
                    FileManagerResponse response = new FileManagerResponse();
                    response.Error = new ErrorDetails { Code = "401", Message = "Restricted to modify the root folder." };
                    return operation.ToCamelCase(response);
                }
            }
            switch (args.Action)
            {
                case "read":
                    // reads the file(s) or folder(s) from the given path.
                    return operation.ToCamelCase(operation.GetFiles(args.Path, args.ShowHiddenItems));
                case "delete":
                    // deletes the selected file(s) or folder(s) from the given path.
                    operation.Response = Response;
                    return operation.ToCamelCase(operation.Delete(args.Path, args.Names));
                case "copy":
                    // copies the selected file(s) or folder(s) from a path and then pastes them into a given target path.
                    return operation.ToCamelCase(operation.Copy(args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData));
                case "move":
                    // cuts the selected file(s) or folder(s) from a path and then pastes them into a given target path.
                    return operation.ToCamelCase(operation.Move(args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData));
                case "details":
                    // gets the details of the selected file(s) or folder(s).
                    return operation.ToCamelCase(operation.Details(args.Path, args.Names, args.Data));
                case "create":
                    // creates a new folder in a given path.
                    return operation.ToCamelCase(operation.Create(args.Path, args.Name));
                case "search":
                    // gets the list of file(s) or folder(s) from a given path based on the searched key string.
                    return operation.ToCamelCase(operation.Search(args.Path, args.SearchString, args.ShowHiddenItems, args.CaseSensitive));
                case "rename":
                    // renames a file or folder.
                    return operation.ToCamelCase(operation.Rename(args.Path, args.Name, args.NewName, false, args.ShowFileExtension, args.Data));
            }
            return null;
        }

        // uploads the file(s) into a specified path
        public IActionResult Upload(string path, IList<IFormFile> uploadFiles, string action)
        {
            FileManagerResponse uploadResponse;
            foreach (var file in uploadFiles)
            {
                var folders = file.FileName.Split('/');
                // checking the folder upload
                if (folders.Length > 1)
                {
                    for (var i = 0; i < folders.Length - 1; i++)
                    {
                        string newDirectoryPath = Path.Combine(basePath + path, folders[i]);
                        if (!Directory.Exists(newDirectoryPath))
                        {
                            operation.ToCamelCase(operation.Create(path, folders[i]));
                        }
                        path += folders[i] + "/";
                    }
                }
            }
            uploadResponse = operation.Upload(path, uploadFiles, action, null);
            if (uploadResponse.Error != null)
            {
                Response.Clear();
                Response.ContentType = "application/json; charset=utf-8";
                Response.StatusCode = Convert.ToInt32(uploadResponse.Error.Code);
                Response.HttpContext.Features.Get<IHttpResponseFeature>().ReasonPhrase = uploadResponse.Error.Message;
            }
            return Content("");
        }

        // downloads the selected file(s) and folder(s)
        public IActionResult Download(string downloadInput)
        {
            FileManagerDirectoryContent args = JsonConvert.DeserializeObject<FileManagerDirectoryContent>(downloadInput);
            return operation.Download(args.Path, args.Names, args.Data);
        }

        // gets the image(s) from the given path
        public IActionResult GetImage(FileManagerDirectoryContent args)
        {
            return operation.GetImage(args.Path, args.Id, false, null, null);
        }
        public AccessDetails GetRules()
        {
            AccessDetails accessDetails = new AccessDetails();
            // List<AccessRule> Rules = new List<AccessRule> {
            //     //Deny writing for particular folder
            //     new AccessRule { Path = "/Documents", Role = "Document Manager", Read = Permission.Allow, Write = Permission.Deny, Copy = Permission.Allow, WriteContents = Permission.Deny, Upload = Permission.Deny, Download = Permission.Deny, IsFile = false },
            //     // Deny writing for particular file
            //     new AccessRule { Path = "/Pictures/Employees/Adam.png", Role = "Document Manager", Read = Permission.Allow, Write = Permission.Deny, Copy = Permission.Deny, Download = Permission.Deny, IsFile = true },
            //     // Folder Rule
            //     new AccessRule { Path = "/", Role = "Document Manager", Read = Permission.Allow, Write = Permission.Deny, Copy = Permission.Deny, WriteContents = Permission.Deny, Upload = Permission.Deny, Download = Permission.Deny, IsFile = false },
            // };
            // accessDetails.AccessRules = Rules;
            // accessDetails.Role = "Document Manager";
            return accessDetails;
        }
    }

}
