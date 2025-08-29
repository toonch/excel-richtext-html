package com.yipt.outbound.dao;
import javax.xml.bind.annotation.*;
import java.util.List;

@XmlRootElement(name = "GUIDMappings")
@XmlAccessorType(XmlAccessType.FIELD)
public class GUIDMappings {
	@XmlElement(name = "Mapping")
    private List<Mapping> mappings;

    // Getters and setters
    public List<Mapping> getMappings() {
        return mappings;
    }

    public void setMappings(List<Mapping> mappings) {
        this.mappings = mappings;
    }

    @XmlAccessorType(XmlAccessType.FIELD)
    public static class Mapping {

        @XmlElement(name = "GUID")
        private String guid;

        @XmlElement(name = "Name")
        private String name;

        // Getters and setters
        public String getGuid() {
            return guid;
        }

        public void setGuid(String guid) {
            this.guid = guid;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }
}
