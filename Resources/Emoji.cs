// Auto-generated. Do not edit by hand.
// 80 document-creation emojis (48x48 PNG), grouped into 8 categories.
//
// Setup options:
//   1) Add Resources/Emojis.resx to your project (it references the PNGs by relative path).
//   2) OR: add each PNG with Build Action = "Embedded Resource" and use the EmojiLoader below.
//
using System.Drawing;
using System.IO;
using System.Reflection;

namespace DocumentEmojis
{
    /// <summary>
    /// Strongly-typed accessors for the document emoji set.
    /// Loads each glyph as a 48x48 <see cref="Bitmap"/> from an embedded resource stream.
    /// Default resource key format: "{AssemblyName}.Resources.Emojis.{Section}.{name}.png".
    /// Override <see cref="ResourceNamespace"/> if your project uses a different default namespace.
    /// </summary>
    public static class Emoji
    {
        /// <summary>Default-namespace prefix used to build embedded-resource keys.</summary>
        public static string ResourceNamespace { get; set; } = "DocumentEmojis.Resources.Emojis";

        private static Bitmap Load(string section, string name)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resName = $"{ResourceNamespace}.{section}.{name}.png";
            using var stream = asm.GetManifestResourceStream(resName)
                ?? throw new FileNotFoundException($"Embedded resource not found: {resName}");
            return new Bitmap(stream);
        }

        public static class Status
        {
            /// <summary>archive</summary>
            public static Bitmap Archive => Load("Status", "archive");
            /// <summary>blocked</summary>
            public static Bitmap Blocked => Load("Status", "blocked");
            /// <summary>done</summary>
            public static Bitmap Done => Load("Status", "done");
            /// <summary>draft</summary>
            public static Bitmap Draft => Load("Status", "draft");
            /// <summary>in progress</summary>
            public static Bitmap InProgress => Load("Status", "in_progress");
            /// <summary>new</summary>
            public static Bitmap New_ => Load("Status", "new");
            /// <summary>on hold</summary>
            public static Bitmap OnHold => Load("Status", "on_hold");
            /// <summary>review</summary>
            public static Bitmap Review => Load("Status", "review");
            /// <summary>todo</summary>
            public static Bitmap Todo => Load("Status", "todo");
            /// <summary>waiting</summary>
            public static Bitmap Waiting => Load("Status", "waiting");
        }

        public static class Reactions
        {
            /// <summary>celebrate</summary>
            public static Bitmap Celebrate => Load("Reactions", "celebrate");
            /// <summary>comment</summary>
            public static Bitmap Comment => Load("Reactions", "comment");
            /// <summary>eyes</summary>
            public static Bitmap Eyes => Load("Reactions", "eyes");
            /// <summary>heart</summary>
            public static Bitmap Heart => Load("Reactions", "heart");
            /// <summary>laugh</summary>
            public static Bitmap Laugh => Load("Reactions", "laugh");
            /// <summary>mind blown</summary>
            public static Bitmap MindBlown => Load("Reactions", "mind_blown");
            /// <summary>raised hand</summary>
            public static Bitmap RaisedHand => Load("Reactions", "raised_hand");
            /// <summary>thinking</summary>
            public static Bitmap Thinking => Load("Reactions", "thinking");
            /// <summary>thumbs down</summary>
            public static Bitmap ThumbsDown => Load("Reactions", "thumbs_down");
            /// <summary>thumbs up</summary>
            public static Bitmap ThumbsUp => Load("Reactions", "thumbs_up");
        }

        public static class Actions
        {
            /// <summary>copy</summary>
            public static Bitmap Copy => Load("Actions", "copy");
            /// <summary>delete</summary>
            public static Bitmap Delete_ => Load("Actions", "delete");
            /// <summary>download</summary>
            public static Bitmap Download => Load("Actions", "download");
            /// <summary>edit</summary>
            public static Bitmap Edit => Load("Actions", "edit");
            /// <summary>link</summary>
            public static Bitmap Link => Load("Actions", "link");
            /// <summary>print</summary>
            public static Bitmap Print => Load("Actions", "print");
            /// <summary>save</summary>
            public static Bitmap Save => Load("Actions", "save");
            /// <summary>search</summary>
            public static Bitmap Search => Load("Actions", "search");
            /// <summary>share</summary>
            public static Bitmap Share => Load("Actions", "share");
            /// <summary>upload</summary>
            public static Bitmap Upload => Load("Actions", "upload");
        }

        public static class Priority
        {
            /// <summary>bookmark</summary>
            public static Bitmap Bookmark => Load("Priority", "bookmark");
            /// <summary>flagged</summary>
            public static Bitmap Flagged => Load("Priority", "flagged");
            /// <summary>high</summary>
            public static Bitmap High => Load("Priority", "high");
            /// <summary>locked</summary>
            public static Bitmap Locked => Load("Priority", "locked");
            /// <summary>low</summary>
            public static Bitmap Low => Load("Priority", "low");
            /// <summary>medium</summary>
            public static Bitmap Medium => Load("Priority", "medium");
            /// <summary>pinned</summary>
            public static Bitmap Pinned => Load("Priority", "pinned");
            /// <summary>star</summary>
            public static Bitmap Star => Load("Priority", "star");
            /// <summary>urgent</summary>
            public static Bitmap Urgent => Load("Priority", "urgent");
            /// <summary>warning</summary>
            public static Bitmap Warning => Load("Priority", "warning");
        }

        public static class People
        {
            /// <summary>assign</summary>
            public static Bitmap Assign => Load("People", "assign");
            /// <summary>collaborators</summary>
            public static Bitmap Collaborators => Load("People", "collaborators");
            /// <summary>group chat</summary>
            public static Bitmap GroupChat => Load("People", "group_chat");
            /// <summary>handover</summary>
            public static Bitmap Handover => Load("People", "handover");
            /// <summary>handshake</summary>
            public static Bitmap Handshake => Load("People", "handshake");
            /// <summary>invite</summary>
            public static Bitmap Invite => Load("People", "invite");
            /// <summary>mention</summary>
            public static Bitmap Mention => Load("People", "mention");
            /// <summary>person</summary>
            public static Bitmap Person => Load("People", "person");
            /// <summary>request</summary>
            public static Bitmap Request => Load("People", "request");
            /// <summary>team</summary>
            public static Bitmap Team => Load("People", "team");
        }

        public static class Tools
        {
            /// <summary>attachment</summary>
            public static Bitmap Attachment => Load("Tools", "attachment");
            /// <summary>calendar</summary>
            public static Bitmap Calendar => Load("Tools", "calendar");
            /// <summary>chart</summary>
            public static Bitmap Chart => Load("Tools", "chart");
            /// <summary>clipboard</summary>
            public static Bitmap Clipboard => Load("Tools", "clipboard");
            /// <summary>document</summary>
            public static Bitmap Document => Load("Tools", "document");
            /// <summary>folder</summary>
            public static Bitmap Folder => Load("Tools", "folder");
            /// <summary>image</summary>
            public static Bitmap Image => Load("Tools", "image");
            /// <summary>pen</summary>
            public static Bitmap Pen => Load("Tools", "pen");
            /// <summary>pencil</summary>
            public static Bitmap Pencil => Load("Tools", "pencil");
            /// <summary>table</summary>
            public static Bitmap Table => Load("Tools", "table");
        }

        public static class Pointers
        {
            /// <summary>callout</summary>
            public static Bitmap Callout => Load("Pointers", "callout");
            /// <summary>down</summary>
            public static Bitmap Down => Load("Pointers", "down");
            /// <summary>expand</summary>
            public static Bitmap Expand => Load("Pointers", "expand");
            /// <summary>left</summary>
            public static Bitmap Left => Load("Pointers", "left");
            /// <summary>point</summary>
            public static Bitmap Point => Load("Pointers", "point");
            /// <summary>redo</summary>
            public static Bitmap Redo => Load("Pointers", "redo");
            /// <summary>refresh</summary>
            public static Bitmap Refresh => Load("Pointers", "refresh");
            /// <summary>right</summary>
            public static Bitmap Right => Load("Pointers", "right");
            /// <summary>undo</summary>
            public static Bitmap Undo => Load("Pointers", "undo");
            /// <summary>up</summary>
            public static Bitmap Up => Load("Pointers", "up");
        }

        public static class Decoration
        {
            /// <summary>award</summary>
            public static Bitmap Award => Load("Decoration", "award");
            /// <summary>confetti</summary>
            public static Bitmap Confetti => Load("Decoration", "confetti");
            /// <summary>crown</summary>
            public static Bitmap Crown => Load("Decoration", "crown");
            /// <summary>fire</summary>
            public static Bitmap Fire => Load("Decoration", "fire");
            /// <summary>gem</summary>
            public static Bitmap Gem => Load("Decoration", "gem");
            /// <summary>idea</summary>
            public static Bitmap Idea => Load("Decoration", "idea");
            /// <summary>rocket</summary>
            public static Bitmap Rocket => Load("Decoration", "rocket");
            /// <summary>sparkle</summary>
            public static Bitmap Sparkle => Load("Decoration", "sparkle");
            /// <summary>target</summary>
            public static Bitmap Target => Load("Decoration", "target");
            /// <summary>trophy</summary>
            public static Bitmap Trophy => Load("Decoration", "trophy");
        }

    }
}
