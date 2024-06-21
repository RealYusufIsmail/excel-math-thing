import groovy.xml.dom.DOMCategory.attributes
import java.util.*

plugins {
    id("java")
    id("com.github.johnrengelman.shadow") version "8.1.1"
    application
}

group = "io.github.realyusufismail"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")
    implementation("org.jfree:jfreechart:1.5.4")
    implementation("org.apache.commons:commons-lang3:3.14.0")

    testImplementation(platform("org.junit:junit-bom:5.10.2"))
    testImplementation("org.junit.jupiter:junit-jupiter")
}

tasks.test {
    useJUnitPlatform()
}

java {
    sourceCompatibility = JavaVersion.VERSION_1_8
    targetCompatibility = JavaVersion.VERSION_1_8
}

application {
    mainClass.set("io.github.realyusufismail.Main")
}

tasks {
    shadowJar {
        archiveBaseName.set("excel-chart-generator")
        manifest {
            attributes(
                "Main-Class" to "io.github.realyusufismail.Main",
                "Implementation-Title" to "Excel Chart Generator",
                "Implementation-Version" to "1.0.0",
                "Built-By" to System.getProperty("user.name"),
                "Built-Date" to Date(),
                "Built-JDK" to System.getProperty("java.version"),
                "Built-Gradle" to gradle.gradleVersion)
        }
    }
}
