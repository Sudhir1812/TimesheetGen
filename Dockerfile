# Use OpenJDK 17 slim as the base image
FROM eclipse-temurin:17-jdk

# Set the app's jar location (Maven puts it in target/)
ARG JAR_FILE=target/TimesheetGen-0.0.1-SNAPSHOT.jar

# Copy jar to container
COPY ${JAR_FILE} app.jar

# Expose port 8082
EXPOSE 8082

# Run Spring Boot app
ENTRYPOINT ["java", "-jar", "/app.jar"]