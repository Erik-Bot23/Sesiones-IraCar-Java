# Etapa de build (compila jar)
FROM maven:3.9-eclipse-temurin-17 AS build
WORKDIR /build
COPY pom.xml .
RUN mvn -q -DskipTests dependency:go-offline
COPY src ./src
RUN mvn -q -DskipTests package

#Etapa de runtime (imagen ligera)
FROM eclipse-temurin:17-jre
WORKDIR /app

#Copia el jar generado
COPY --from=build /build/target/agendadesesiones-*.jar app.jar
EXPOSE 8081

#Usa un usuario no root por buenas prácticas
RUN useradd -m appuser
USER appuser
ENTRYPOINT [ "java", "-XX:+UseContainerSupport", "-XX:MaxRAMPercentage=75", "-jar", "app.jar" ]