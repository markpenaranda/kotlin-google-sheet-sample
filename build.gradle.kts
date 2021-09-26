import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
    kotlin("jvm") version "1.5.31"
    application
}

group = "com.markpenaranda"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {

    implementation("com.google.api-client:google-api-client:1.30.4")
    implementation("com.google.oauth-client:google-oauth-client-jetty:1.30.6")
    implementation("com.google.apis:google-api-services-sheets:v4-rev581-1.25.0")
    implementation("com.google.apis:google-api-services-drive:v3-rev110-1.23.0")

    testImplementation(kotlin("test"))
}

tasks.test {
    useJUnit()
}

tasks.withType<KotlinCompile>() {
    kotlinOptions.jvmTarget = "1.8"
}

application {
    mainClass.set("MainKt")
}
