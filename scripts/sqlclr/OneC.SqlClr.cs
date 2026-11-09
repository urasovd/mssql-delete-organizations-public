using System;
using System.Collections;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SqlServer.Server;

namespace OneC.SqlClr
{
    public static class OneCMetadata
    {
        private static readonly Regex DbNamesEntry = new Regex(
            @"\{(?<id>[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}),""(?<token>[^""]+)"",(?<code>[0-9]+)\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ConfigNameEntry = new Regex(
            @"\{[01],0,(?<id>[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})\},""(?<name>[^""]*)""",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex TableHeader = new Regex(
            // Tabular parts carry their owner in the fourth position, e.g.
            // {"VT4415","I",0,"Document175",{25,...  -> Document175_VT4415.
            // Requiring an empty owner here used to drop every tabular part.
            @"\{""(?<table>[A-Za-z][A-Za-z0-9_]*)"",""[NI]"",[0-9]+,""(?<owner>[A-Za-z0-9_]*)"",\s*(?<fieldsOpen>\{)[0-9]+,",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex FixedReferenceField = new Regex(
            @"\{""(?<field>[A-Za-z][A-Za-z0-9_]*)"",[01],\s*\{[0-9]+,\s*\{""R"",0,0,""(?<target>[A-Za-z][A-Za-z0-9_]*)"",3\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex FieldHeader = new Regex(
            @"\{""(?<field>[A-Za-z][A-Za-z0-9_]*)"",[01],\s*(?<typesOpen>\{)(?<count>[0-9]+),",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex MetadataReference = new Regex(
            @"\{1,(?<id>[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ReferenceType = new Regex(
            @"\{""R"",0,0,""(?<target>[A-Za-z0-9_]*)"",[0-9]+\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        [SqlFunction(DataAccess = DataAccessKind.None, IsDeterministic = true, IsPrecise = true)]
        [return: SqlFacet(MaxSize = -1)]
        public static SqlString InflateUtf8(SqlBytes input)
        {
            if (input == null || input.IsNull)
                return SqlString.Null;

            return new SqlString(Encoding.UTF8.GetString(Inflate(input.Value)));
        }

        [SqlFunction(DataAccess = DataAccessKind.None, IsDeterministic = true, IsPrecise = true)]
        [return: SqlFacet(MaxSize = -1)]
        public static SqlString DecodeUtf8(SqlBytes input)
        {
            if (input == null || input.IsNull)
                return SqlString.Null;

            return new SqlString(Encoding.UTF8.GetString(input.Value));
        }

        [SqlFunction(
            DataAccess = DataAccessKind.None,
            IsDeterministic = true,
            IsPrecise = true,
            FillRowMethodName = "FillDbNameRow",
            TableDefinition = "MetadataId uniqueidentifier, Token nvarchar(64), Code int")]
        public static IEnumerable ParseDbNames(SqlBytes input)
        {
            if (input == null || input.IsNull)
                yield break;

            string contents = Encoding.UTF8.GetString(Inflate(input.Value));
            foreach (Match match in DbNamesEntry.Matches(contents))
            {
                yield return new DbNameRow(
                    new Guid(match.Groups["id"].Value),
                    match.Groups["token"].Value,
                    Int32.Parse(match.Groups["code"].Value, CultureInfo.InvariantCulture));
            }
        }

        public static void FillDbNameRow(object value, out SqlGuid metadataId, out SqlString token, out SqlInt32 code)
        {
            DbNameRow row = (DbNameRow)value;
            metadataId = new SqlGuid(row.MetadataId);
            token = new SqlString(row.Token);
            code = new SqlInt32(row.Code);
        }

        [SqlFunction(
            DataAccess = DataAccessKind.None,
            IsDeterministic = true,
            IsPrecise = true,
            FillRowMethodName = "FillConfigNameRow",
            TableDefinition = "MetadataId uniqueidentifier, LogicalName nvarchar(256)")]
        public static IEnumerable ParseConfigNames(SqlBytes input)
        {
            if (input == null || input.IsNull)
                yield break;

            string contents = Encoding.UTF8.GetString(Inflate(input.Value));
            foreach (Match match in ConfigNameEntry.Matches(contents))
            {
                yield return new ConfigNameRow(
                    new Guid(match.Groups["id"].Value),
                    match.Groups["name"].Value);
            }
        }

        public static void FillConfigNameRow(object value, out SqlGuid metadataId, out SqlString logicalName)
        {
            ConfigNameRow row = (ConfigNameRow)value;
            metadataId = new SqlGuid(row.MetadataId);
            logicalName = new SqlString(row.LogicalName);
        }

        [SqlFunction(
            DataAccess = DataAccessKind.None,
            IsDeterministic = true,
            IsPrecise = true,
            FillRowMethodName = "FillReferenceEdgeRow",
            TableDefinition = "SourceTable nvarchar(128), SourceField nvarchar(128), TargetTable nvarchar(128)")]
        public static IEnumerable ParseFixedReferenceEdges(SqlBytes schema, SqlString targetTable)
        {
            if (schema == null || schema.IsNull)
                yield break;

            string target = targetTable == null || targetTable.IsNull ? String.Empty : targetTable.Value;
            string contents = Encoding.UTF8.GetString(schema.Value);

            foreach (Match tableMatch in TableHeader.Matches(contents))
            {
                int fieldsOpen = tableMatch.Groups["fieldsOpen"].Index;
                int fieldsClose = FindMatchingBrace(contents, fieldsOpen);
                if (fieldsClose < 0)
                    continue;

                string fields = contents.Substring(fieldsOpen, fieldsClose - fieldsOpen + 1);
                string owner = tableMatch.Groups["owner"].Value;
                string sourceTable = owner.Length == 0
                    ? tableMatch.Groups["table"].Value
                    : owner + "_" + tableMatch.Groups["table"].Value;
                foreach (Match fieldMatch in FixedReferenceField.Matches(fields))
                {
                    string edgeTarget = fieldMatch.Groups["target"].Value;
                    if (target.Length != 0 && !String.Equals(target, edgeTarget, StringComparison.OrdinalIgnoreCase))
                        continue;

                    yield return new ReferenceEdgeRow(
                        sourceTable,
                        fieldMatch.Groups["field"].Value,
                        edgeTarget);
                }
            }
        }

        // Every reference type of every field, including composite fields.
        // TypeCount is the number of type entries of the field: 1 means a plain
        // fixed reference; more means a composite field.  An empty TargetTable
        // stands for "any reference" and always has an _RTRef column.
        public static IEnumerable ParseReferenceFieldTargets(SqlBytes schema)
        {
            if (schema == null || schema.IsNull)
                yield break;

            string contents = Encoding.UTF8.GetString(schema.Value);
            foreach (Match tableMatch in TableHeader.Matches(contents))
            {
                int fieldsOpen = tableMatch.Groups["fieldsOpen"].Index;
                int fieldsClose = FindMatchingBrace(contents, fieldsOpen);
                if (fieldsClose < 0)
                    continue;

                string owner = tableMatch.Groups["owner"].Value;
                string sourceTable = owner.Length == 0
                    ? tableMatch.Groups["table"].Value
                    : owner + "_" + tableMatch.Groups["table"].Value;
                string fields = contents.Substring(fieldsOpen, fieldsClose - fieldsOpen + 1);
                foreach (Match fieldMatch in FieldHeader.Matches(fields))
                {
                    int typesOpen = fieldMatch.Groups["typesOpen"].Index;
                    int typesClose = FindMatchingBrace(fields, typesOpen);
                    if (typesClose < 0)
                        continue;

                    int typeCount = Int32.Parse(fieldMatch.Groups["count"].Value, CultureInfo.InvariantCulture);
                    string types = fields.Substring(typesOpen, typesClose - typesOpen + 1);
                    foreach (Match typeMatch in ReferenceType.Matches(types))
                    {
                        yield return new ReferenceTargetRow(
                            sourceTable,
                            fieldMatch.Groups["field"].Value,
                            typeMatch.Groups["target"].Value,
                            typeCount);
                    }
                }
            }
        }

        // Owner list of a catalog ("Владельцы" in the designer).  In the
        // compressed Config file of the catalog it is the first {0,N,...}
        // block right after the name block {1,0,<metadataId>},"Name",...;
        // every item is {"#",157fa490-...,{1,<owner metadata id>}}.  The same
        // item marker is used elsewhere (input-on-basis lists, etc.), so the
        // position, not the marker, identifies the owner list.  Returns null
        // when the block is not where expected.
        public static string[] ParseCatalogOwners(SqlBytes configFile, string metadataId)
        {
            if (configFile == null || configFile.IsNull)
                return null;

            string contents = Encoding.UTF8.GetString(Inflate(configFile.Value));
            Regex ownerBlock = new Regex(
                @"\{1,0," + Regex.Escape(metadataId) + @"\},""[^""]*"",[^\r\n]*\r?\n?(?:[^\r\n]*\r?\n)*?\},[0-9]+,[0-9]+,\s*\{0,(?<count>[0-9]+)(?<items>(?:,\s*\{""#"",157fa490-4ce9-11d4-9415-008048da11f9,\s*\{1,[0-9a-fA-F-]{36}\}\s*\})*)\s*\}",
                RegexOptions.CultureInvariant);
            Match match = ownerBlock.Match(contents);
            if (!match.Success)
                return null;

            MatchCollection items = MetadataReference.Matches(match.Groups["items"].Value);
            if (items.Count != Int32.Parse(match.Groups["count"].Value, CultureInfo.InvariantCulture))
                return null;

            string[] owners = new string[items.Count];
            for (int index = 0; index < items.Count; index++)
                owners[index] = items[index].Groups["id"].Value.ToLowerInvariant();
            return owners;
        }

        public static void FillReferenceEdgeRow(object value, out SqlString sourceTable, out SqlString sourceField, out SqlString targetTable)
        {
            ReferenceEdgeRow row = (ReferenceEdgeRow)value;
            sourceTable = new SqlString(row.SourceTable);
            sourceField = new SqlString(row.SourceField);
            targetTable = new SqlString(row.TargetTable);
        }

        private static byte[] Inflate(byte[] input)
        {
            using (MemoryStream source = new MemoryStream(input, false))
            using (DeflateStream inflater = new DeflateStream(source, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                inflater.CopyTo(output);
                return output.ToArray();
            }
        }

        private static int FindMatchingBrace(string contents, int openIndex)
        {
            int depth = 0;
            bool inString = false;

            for (int index = openIndex; index < contents.Length; index++)
            {
                char current = contents[index];
                if (current == '"')
                {
                    if (inString && index + 1 < contents.Length && contents[index + 1] == '"')
                    {
                        index++;
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (inString)
                    continue;

                if (current == '{')
                    depth++;
                else if (current == '}')
                {
                    depth--;
                    if (depth == 0)
                        return index;
                }
            }

            return -1;
        }

        private sealed class DbNameRow
        {
            public readonly Guid MetadataId;
            public readonly string Token;
            public readonly int Code;
            public DbNameRow(Guid metadataId, string token, int code) { MetadataId = metadataId; Token = token; Code = code; }
        }

        private sealed class ConfigNameRow
        {
            public readonly Guid MetadataId;
            public readonly string LogicalName;
            public ConfigNameRow(Guid metadataId, string logicalName) { MetadataId = metadataId; LogicalName = logicalName; }
        }

        public sealed class ReferenceTargetRow
        {
            public readonly string SourceTable;
            public readonly string SourceField;
            public readonly string TargetTable;
            public readonly int TypeCount;
            public ReferenceTargetRow(string sourceTable, string sourceField, string targetTable, int typeCount)
            {
                SourceTable = sourceTable;
                SourceField = sourceField;
                TargetTable = targetTable;
                TypeCount = typeCount;
            }
        }

        private sealed class ReferenceEdgeRow
        {
            public readonly string SourceTable;
            public readonly string SourceField;
            public readonly string TargetTable;
            public ReferenceEdgeRow(string sourceTable, string sourceField, string targetTable)
            {
                SourceTable = sourceTable;
                SourceField = sourceField;
                TargetTable = targetTable;
            }
        }
    }
}
