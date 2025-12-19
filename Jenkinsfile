pipeline {
    agent any

    environment {
        IMAGE_NAME = "timesheetgen:1.0"
    }

    stages {
        stage('Checkout') {
            steps {
                git branch: 'main', url: 'https://github.com/Sudhir1812/TimesheetGen.git'
            }
        }

        stage('Build') {
            steps {
                // Use bat for Windows
                bat 'mvnw.cmd clean package'
            }
        }

        stage('Docker Build') {
            steps {
                bat "docker build -t %IMAGE_NAME% ."
            }
        }

        stage('Deploy to Kubernetes') {
            steps {
                bat 'kubectl apply -f k8s\\'
            }
        }
    }
}
