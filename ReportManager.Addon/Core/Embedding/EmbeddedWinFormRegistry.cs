using ReportManager.Addon.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core.Embedding
{
    public sealed class EmbeddedWinFormRegistry
    {
        private readonly Dictionary<string, EmbeddedWinFormHost> _hosts = new Dictionary<string, EmbeddedWinFormHost>(StringComparer.OrdinalIgnoreCase);

        public EmbeddedWinFormHost GetOrCreate(SAPbouiCOM.Application app, Logger log, string hostSapFormUid, string hostTitle)
        {
            if (!_hosts.TryGetValue(hostSapFormUid, out var host))
            {
                host = new EmbeddedWinFormHost(app, log, hostSapFormUid, hostTitle);
                _hosts[hostSapFormUid] = host;
            }
            return host;
        }

        public void DisposeHost(string hostSapFormUid)
        {
            if (string.IsNullOrWhiteSpace(hostSapFormUid)) return;

            if (_hosts.TryGetValue(hostSapFormUid, out var host))
            {
                try { host.Dispose(); } catch { }
                _hosts.Remove(hostSapFormUid);
            }
        }

        public void DisposeAll()
        {
            foreach (var kv in _hosts)
            {
                try { kv.Value.Dispose(); } catch { }
            }
            _hosts.Clear();
        }
    }
}
