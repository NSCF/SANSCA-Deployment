# Configuration details

## Basic information

**Your LA Portal Long Name:** South African Natural Science Collections Atlas  
**Short Name:** SANSCA  
**Use SSL?:** Yes  
**The domain of your LA Portal:** sansca.org.za  
> **_NOTE:_** DNS A record information will be reuqired to setup sansca.org.za

**Portal branding:** clean  

## Location information

**Short Name of Map Area:** South Africa
## Servers
## Services

**Minimum Viable Install:**  
*See https://github.com/AtlasOfLivingAustralia/documentation/wiki/Infrastructure-Requirements#core-components-for-a-living-atlas* for additional information.
* branding
* cas
* biocache_backend [Apache Cassandra] (Core)
* Apache Solr (Core)
* collectory (Core)
* ala_hub (Core)
* biocache_services (Core)
* ala_bie
* bie_index
* nameindexer
* image_service (Core)
* biocache_cli (Core)

**Full Instalation:**
* Add *lists, regions, logger, alerts, doi, dashboard, sds (sensitive species management), data_quality, spatial*

**Deployment Sequence:**  
* branding
* cas
* biocache_backend
* solr
* collectory
* ala_hub
* biocache_service
* ala_bie
* bie_index
* lists
* regions
* logger
* nameindexer
* images
* alerts
* biocache_cli
* doi
* dashboard
* sds
* data_quality
* spatial

## Which services on which servers
## Server Information

**Default user in your servers:** ???  
**Advanced SSH options:** ???
