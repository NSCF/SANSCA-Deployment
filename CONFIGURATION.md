# Configuration details

## Basic information

**Your LA Portal Long Name:** South African Natural Science Collections Atlas  
**Short Name:** SANSCA  
**Use SSL?:** Yes  
**The domain of your LA Portal:** sansca.org.za  
> **_NOTE:_** DNS A record information will be required to setup sansca.org.za

**Portal branding:** clean  

## Location information

**Short Name of Map Area:** South Africa
## Servers
## Services
See https://github.com/AtlasOfLivingAustralia/documentation/wiki/Infrastructure-Requirements#core-components-for-a-living-atlas* for additional information.

**Core Install**
> **_NOTE:_** The core installation modules can be used for the first test deployment. The biochache_backend consists of the biocache-store and biocache_cli modules. 
* biocache-store (biocache_backend) [Apache Cassandra]
* solr [Apache]
* collectory
* records (ala_hub)
* records-ws (biocache_services)
* image_service
* biocache_cli (biocache_backend)

**Minimum Viable SANSCA Installation:**
> **_NOTE:_** The following modules are the minimum required to launch SANSCA.
* branding
* auth (cas)
* biocache-store (biocache_backend)
* solr
* collectory
* records (ala_hub)
* records-ws (biocache_services)
* species (ala_bie)
* species-ws (bie_index)
* nameindexer
* image_service
* biocache_cli (biocache_backend)

**Full Installation:**
> **_NOTE:_** To maximise the potential of the ALA platform, a full installation is the ultimate goal.
* Add *lists, regions, logger, alerts, doi, dashboard, sds (sensitive species management), data_quality, spatial*

**Full Deployment Sequence:**  
* branding
* auth (cas)
* biocache-store (biocache_backend)
* solr
* collectory
* records (ala_hub)
* records-ws (biocache_services)
* species (ala_bie)
* species-ws (bie_index)
* lists
* regions
* logger
* nameindexer
* images
* alerts
* biocache_cli (biocache_backend)
* doi
* dashboard
* sds
* data_quality
* spatial

## Which services on which servers
## Server Information

**Default user in your servers:** ???  
**Advanced SSH options:** ???
