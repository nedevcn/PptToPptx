using System;
using System.IO;
using System.Collections.Generic;

namespace Nedev.FileConverters.PptToPptx
{
    public class OleCompoundFile
    {
        private readonly Stream _stream;
        private readonly byte[] _header = new byte[512];
        private int _sectorSize;
        private int _miniSectorSize;
        private int _fatSectorCount;
        private int _firstDirectorySector;
        private int _firstMiniFatSector;
        private int _miniFatSectorCount;
        private int _miniStreamCutoff;
        private int _firstDifatSector;
        private int _difatSectorCount;
        private List<int> _fat = new List<int>();
        private List<int> _miniFat = new List<int>();
        private byte[] _miniStreamData;
        public IReadOnlyList<DirectoryEntry> DirectoryList => _directoryList;
        private readonly List<DirectoryEntry> _directoryList = new List<DirectoryEntry>(); // non-empty entries for enumeration
        private readonly List<DirectoryEntry> _directoryBySid = new List<DirectoryEntry>(); // includes empties to preserve SID indexing
        private readonly Dictionary<string, List<int>> _nameToSids = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        
        private const int END_OF_CHAIN = -2;       // 0xFFFFFFFE
        private const int FREE_SECT = -1;           // 0xFFFFFFFF
        
        public OleCompoundFile(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            ReadHeader();
        }
        
        private void ReadHeader()
        {
            _stream.Position = 0;
            _stream.Read(_header, 0, 512);
            
            // 验证文件头 magic bytes
            byte[] expectedHeader = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            for (int i = 0; i < 8; i++)
            {
                if (_header[i] != expectedHeader[i])
                {
                    throw new InvalidDataException("Not a valid OLE Compound File");
                }
            }
            
            // 读取扇区大小 (offset 30, 2 bytes) — 通常 9 (512) 或 12 (4096)
            int sectorShift = BitConverter.ToUInt16(_header, 30);
            _sectorSize = 1 << sectorShift;
            
            // 读取迷你扇区大小 (offset 32, 2 bytes) — 通常 6 (64)
            int miniSectorShift = BitConverter.ToUInt16(_header, 32);
            _miniSectorSize = 1 << miniSectorShift;
            
            // 读取 FAT 扇区数量 (offset 44, 4 bytes)
            _fatSectorCount = BitConverter.ToInt32(_header, 44);
            
            // 读取第一个目录扇区 (offset 48, 4 bytes)
            _firstDirectorySector = BitConverter.ToInt32(_header, 48);
            
            // 迷你流的大小截止 (offset 56, 4 bytes) — 通常 4096
            _miniStreamCutoff = BitConverter.ToInt32(_header, 56);
            
            // 读取第一个 Mini FAT 扇区 (offset 60, 4 bytes)
            _firstMiniFatSector = BitConverter.ToInt32(_header, 60);
            
            // 读取 Mini FAT 扇区数量 (offset 64, 4 bytes)
            _miniFatSectorCount = BitConverter.ToInt32(_header, 64);
            
            // 读取第一个 DIFAT 扇区 (offset 68, 4 bytes)
            _firstDifatSector = BitConverter.ToInt32(_header, 68);
            
            // 读取 DIFAT 扇区数量 (offset 72, 4 bytes)
            _difatSectorCount = BitConverter.ToInt32(_header, 72);
        }
        
        /// <summary>
        /// 将扇区号转为文件中偏移量（扇区0从header之后开始）
        /// </summary>
        private long SectorOffset(int sectorId)
        {
            return (long)(sectorId + 1) * _sectorSize;
        }
        
        public void Parse()
        {
            ReadFat();
            ReadDirectoryEntries();
            ReadMiniFat();
            BuildMiniStream();
        }
        
        private void ReadFat()
        {
            // 先从 header 中的 DIFAT 数组读取 FAT 扇区 (最多 109 个 FAT 扇区)
            List<int> fatSectors = new List<int>();
            for (int i = 0; i < 109 && i < _fatSectorCount; i++)
            {
                int fatSectorId = BitConverter.ToInt32(_header, 76 + i * 4);
                if (fatSectorId == END_OF_CHAIN || fatSectorId == FREE_SECT)
                    break;
                fatSectors.Add(fatSectorId);
            }
            
            // 如果有 DIFAT 扇区链（大文件），继续读取
            if (_difatSectorCount > 0 && _firstDifatSector != END_OF_CHAIN)
            {
                int difatSector = _firstDifatSector;
                for (int d = 0; d < _difatSectorCount && difatSector != END_OF_CHAIN && difatSector != FREE_SECT; d++)
                {
                    byte[] difatData = ReadSectorData(difatSector);
                    int entriesPerSector = _sectorSize / 4 - 1; // 最后一个是下一个 DIFAT 扇区号
                    for (int j = 0; j < entriesPerSector; j++)
                    {
                        int sid = BitConverter.ToInt32(difatData, j * 4);
                        if (sid == END_OF_CHAIN || sid == FREE_SECT)
                            break;
                        fatSectors.Add(sid);
                    }
                    difatSector = BitConverter.ToInt32(difatData, entriesPerSector * 4);
                }
            }
            
            // 读取所有 FAT 扇区内容
            foreach (int fatSectorId in fatSectors)
            {
                byte[] fatSectorData = ReadSectorData(fatSectorId);
                for (int j = 0; j < _sectorSize / 4; j++)
                {
                    int entry = BitConverter.ToInt32(fatSectorData, j * 4);
                    _fat.Add(entry);
                }
            }
        }
        
        private void ReadMiniFat()
        {
            if (_firstMiniFatSector == END_OF_CHAIN || _firstMiniFatSector == FREE_SECT)
                return;
                
            int currentSector = _firstMiniFatSector;
            while (currentSector != END_OF_CHAIN && currentSector != FREE_SECT)
            {
                byte[] data = ReadSectorData(currentSector);
                for (int j = 0; j < _sectorSize / 4; j++)
                {
                    _miniFat.Add(BitConverter.ToInt32(data, j * 4));
                }
                currentSector = GetNextSector(currentSector);
            }
        }
        
        private void BuildMiniStream()
        {
            // Mini Stream 存在 Root Entry 的流中
            if (_directoryBySid.Count == 0) return;

            DirectoryEntry rootEntry = null;
            for (int i = 0; i < _directoryBySid.Count; i++)
            {
                if (_directoryBySid[i].Type == DirectoryEntryType.RootStorage)
                {
                    rootEntry = _directoryBySid[i];
                    break;
                }
            }
            rootEntry ??= _directoryBySid[0];

            if (rootEntry.StartSector == END_OF_CHAIN || rootEntry.StartSector == FREE_SECT)
                return;
                
            _miniStreamData = ReadStreamData(rootEntry.StartSector, rootEntry.Size);
        }
        
        private byte[] ReadSectorData(int sectorId)
        {
            byte[] data = new byte[_sectorSize];
            _stream.Position = SectorOffset(sectorId);
            _stream.Read(data, 0, _sectorSize);
            return data;
        }
        
        private int GetNextSector(int currentSector)
        {
            if (currentSector < 0 || currentSector >= _fat.Count)
                return END_OF_CHAIN;
            return _fat[currentSector];
        }
        
        private void ReadDirectoryEntries()
        {
            int currentSector = _firstDirectorySector;
            int sid = 0;
            while (currentSector != END_OF_CHAIN && currentSector != FREE_SECT)
            {
                byte[] directorySectorData = ReadSectorData(currentSector);
                
                // 每个目录项 128 字节
                for (int i = 0; i < _sectorSize / 128; i++)
                {
                    byte[] entryData = new byte[128];
                    Array.Copy(directorySectorData, i * 128, entryData, 0, 128);
                    
                    DirectoryEntry entry = new DirectoryEntry(entryData, sid);
                    _directoryBySid.Add(entry);

                    if (entry.Type != DirectoryEntryType.Empty && entry.Name.Length > 0)
                    {
                        _directoryList.Add(entry);
                        if (!_nameToSids.TryGetValue(entry.Name, out var list))
                        {
                            list = new List<int>();
                            _nameToSids[entry.Name] = list;
                        }
                        list.Add(entry.Sid);
                    }

                    sid++;
                }
                
                currentSector = GetNextSector(currentSector);
            }
        }
        
        public Stream GetStream(string streamName)
        {
            if (_nameToSids.TryGetValue(streamName, out var sids))
            {
                // Prefer a user stream when names collide.
                foreach (var sid in sids)
                {
                    var e = GetEntryBySid(sid);
                    if (e != null && e.Type == DirectoryEntryType.UserStream)
                        return GetStream(e);
                }
                var entry = GetEntryBySid(sids[0]);
                if (entry != null)
                    return GetStream(entry);
            }
            return null;
        }
        
        public Stream GetStream(DirectoryEntry entry)
        {
            if (entry != null && entry.Type == DirectoryEntryType.UserStream)
            {
                byte[] data;
                if (entry.Size < _miniStreamCutoff && _miniStreamData != null)
                {
                    data = ReadMiniStreamData(entry.StartSector, entry.Size);
                }
                else
                {
                    data = ReadStreamData(entry.StartSector, entry.Size);
                }
                return new MemoryStream(data);
            }
            return null;
        }

        public List<DirectoryEntry> GetEntriesByName(string name)
        {
            var result = new List<DirectoryEntry>();
            foreach (var entry in _directoryList)
            {
                if (string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(entry);
                }
            }
            return result;
        }
        
        /// <summary>
        /// Find a child stream by name inside a storage entry (CFB directory tree).
        /// </summary>
        public Stream GetChildStream(string storageName, string streamName)
        {
            var storage = FindStorageEntry(storageName);
            if (storage == null) return null;

            var child = FindDirectChildByName(storage, streamName, DirectoryEntryType.UserStream);
            if (child != null)
                return GetStream(child);

            return null;
        }

        /// <summary>
        /// Return direct child entries of a storage/root (streams + storages).
        /// </summary>
        public List<DirectoryEntry> GetChildEntries(string storageName)
        {
            var storage = FindStorageEntry(storageName);
            if (storage == null) return new List<DirectoryEntry>();
            return GetChildEntries(storage);
        }

        public List<DirectoryEntry> GetChildEntries(DirectoryEntry storageEntry)
        {
            var result = new List<DirectoryEntry>();
            if (storageEntry == null) return result;
            if (storageEntry.Type != DirectoryEntryType.UserStorage && storageEntry.Type != DirectoryEntryType.RootStorage)
                return result;

            foreach (var child in EnumerateSiblingTree(storageEntry.ChildSid))
            {
                if (child.Type != DirectoryEntryType.Empty && child.Name.Length > 0)
                    result.Add(child);
            }
            return result;
        }

        /// <summary>
        /// Get all storage entries whose names start with a given prefix.
        /// </summary>
        public List<DirectoryEntry> GetStoragesByPrefix(string prefix)
        {
            var result = new List<DirectoryEntry>();
            foreach (var entry in _directoryList)
            {
                if (entry.Type == DirectoryEntryType.UserStorage && entry.Name.StartsWith(prefix))
                {
                    result.Add(entry);
                }
            }
            return result;
        }

        public List<string> GetAllStreamNames()
        {
            var names = new List<string>();
            foreach (var entry in _directoryList)
            {
                names.Add($"[{entry.Type}] {entry.Name} (sector={entry.StartSector}, size={entry.Size})");
            }
            return names;
        }

        private DirectoryEntry GetEntryBySid(int sid)
        {
            if (sid < 0 || sid >= _directoryBySid.Count) return null;
            return _directoryBySid[sid];
        }

        private DirectoryEntry FindStorageEntry(string storageName)
        {
            if (string.IsNullOrEmpty(storageName)) return null;

            if (_nameToSids.TryGetValue(storageName, out var sids))
            {
                foreach (var sid in sids)
                {
                    var e = GetEntryBySid(sid);
                    if (e != null && (e.Type == DirectoryEntryType.UserStorage || e.Type == DirectoryEntryType.RootStorage))
                        return e;
                }
            }

            // Fallback: linear scan (case-insensitive)
            foreach (var entry in _directoryList)
            {
                if ((entry.Type == DirectoryEntryType.UserStorage || entry.Type == DirectoryEntryType.RootStorage) &&
                    string.Equals(entry.Name, storageName, StringComparison.OrdinalIgnoreCase))
                    return entry;
            }

            return null;
        }

        private DirectoryEntry FindDirectChildByName(DirectoryEntry storageEntry, string childName, DirectoryEntryType requiredType)
        {
            if (storageEntry == null || string.IsNullOrEmpty(childName)) return null;

            foreach (var child in EnumerateSiblingTree(storageEntry.ChildSid))
            {
                if (child.Type == requiredType && string.Equals(child.Name, childName, StringComparison.OrdinalIgnoreCase))
                    return child;
            }
            return null;
        }

        /// <summary>
        /// Enumerate nodes in a directory sibling red-black tree (left/right only).
        /// This is used to enumerate direct children of a storage whose ChildSid points to the tree root.
        /// </summary>
        private IEnumerable<DirectoryEntry> EnumerateSiblingTree(int rootSid)
        {
            if (rootSid < 0) yield break;

            var stack = new Stack<int>();
            var visited = new HashSet<int>();
            stack.Push(rootSid);

            while (stack.Count > 0)
            {
                int sid = stack.Pop();
                if (sid < 0 || sid >= _directoryBySid.Count) continue;
                if (!visited.Add(sid)) continue;

                var entry = _directoryBySid[sid];
                yield return entry;

                if (entry.LeftSiblingSid >= 0) stack.Push(entry.LeftSiblingSid);
                if (entry.RightSiblingSid >= 0) stack.Push(entry.RightSiblingSid);
            }
        }
        
        private byte[] ReadStreamData(int startSector, long size)
        {
            var ms = new MemoryStream();
            int currentSector = startSector;
            long remaining = size;
            
            while (currentSector != END_OF_CHAIN && currentSector != FREE_SECT && remaining > 0)
            {
                _stream.Position = SectorOffset(currentSector);
                int bytesToRead = (int)Math.Min(_sectorSize, remaining);
                byte[] buffer = new byte[bytesToRead];
                _stream.Read(buffer, 0, bytesToRead);
                ms.Write(buffer, 0, bytesToRead);
                remaining -= bytesToRead;
                currentSector = GetNextSector(currentSector);
            }
            
            return ms.ToArray();
        }
        
        private byte[] ReadMiniStreamData(int startMiniSector, long size)
        {
            var ms = new MemoryStream();
            int currentMiniSector = startMiniSector;
            long remaining = size;
            
            while (currentMiniSector != END_OF_CHAIN && currentMiniSector != FREE_SECT && remaining > 0)
            {
                long miniOffset = (long)currentMiniSector * _miniSectorSize;
                int bytesToRead = (int)Math.Min(_miniSectorSize, remaining);
                
                if (miniOffset + bytesToRead <= _miniStreamData.Length)
                {
                    ms.Write(_miniStreamData, (int)miniOffset, bytesToRead);
                }
                
                remaining -= bytesToRead;
                
                if (currentMiniSector < _miniFat.Count)
                    currentMiniSector = _miniFat[currentMiniSector];
                else
                    break;
            }
            
            return ms.ToArray();
        }
        
        public enum DirectoryEntryType
        {
            Empty = 0,
            UserStorage = 1,
            UserStream = 2,
            LockBytes = 3,
            Property = 4,
            RootStorage = 5
        }
        
        public class DirectoryEntry
        {
            public int Sid { get; private set; }
            public string Name { get; private set; }
            public DirectoryEntryType Type { get; private set; }
            public int LeftSiblingSid { get; private set; }
            public int RightSiblingSid { get; private set; }
            public int ChildSid { get; private set; }
            public int StartSector { get; private set; }
            public long Size { get; private set; }
            
            public DirectoryEntry(byte[] data, int sid)
            {
                Sid = sid;
                // 名称长度 (offset 64, 2 bytes) — Unicode 字符数 including null
                int nameLen = BitConverter.ToUInt16(data, 64);
                if (nameLen > 2)
                {
                    char[] nameChars = new char[(nameLen - 2) / 2];
                    for (int i = 0; i < nameChars.Length; i++)
                    {
                        nameChars[i] = BitConverter.ToChar(data, i * 2);
                    }
                    Name = new string(nameChars);
                }
                else
                {
                    Name = "";
                }
                
                // 类型 (offset 66, 1 byte)
                Type = (DirectoryEntryType)(data[66] & 0xFF);

                // Sibling/child pointers (CFB directory red-black tree)
                LeftSiblingSid = BitConverter.ToInt32(data, 68);
                RightSiblingSid = BitConverter.ToInt32(data, 72);
                ChildSid = BitConverter.ToInt32(data, 76);
                
                // 起始扇区 (offset 116, 4 bytes)
                StartSector = BitConverter.ToInt32(data, 116);
                
                // 大小 (offset 120, 8 bytes) — v3 uses low 32 bits, v4 uses 64-bit
                try
                {
                    Size = (long)BitConverter.ToUInt64(data, 120);
                }
                catch
                {
                    Size = BitConverter.ToUInt32(data, 120);
                }
            }
        }
    }
}
