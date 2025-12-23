pipeline {
    agent any

    environment {
        APP_NAME   = "timesheetgen"
        IMAGE_TAG  = "${BRANCH_NAME}-${BUILD_NUMBER}"
        IMAGE_NAME = "${APP_NAME}:${IMAGE_TAG}"
        KUBECONFIG = "C:\\ProgramData\\Jenkins\\.kube\\config"
    }

    options {
        disableConcurrentBuilds()
    }

    stages {

        stage('CI - Build & Test') {
            steps {
                bat 'mvnw.cmd clean test'
            }
        }

        stage('Package') {
            steps {
                bat 'mvnw.cmd clean package -DskipTests'
            }
        }

        stage('Docker Build') {
            when {
                anyOf {
                    branch 'develop'
                    branch 'main'
                }
            }
            steps {
                echo "Building Docker image: ${IMAGE_NAME}"
                bat "docker build -t %IMAGE_NAME% ."
            }
        }

        stage('Deploy to TEST') {
            when {
                branch 'develop'
            }
            steps {
                echo 'Deploying to TEST environment'

                // Replace IMAGE_TAG in TEST deployment manifest
                bat """
                powershell -Command "(Get-Content k8s/test/deployment.yaml) `
                -replace 'IMAGE_TAG', '${IMAGE_TAG}' |
                Set-Content k8s/test/deployment.yaml"
                """

                // Create namespace if not exists
                bat 'kubectl get namespace test || kubectl create namespace test'

                // Apply manifests
                bat 'kubectl apply -f k8s/test/ --validate=false'
            }
        }

        stage('Deploy to PROD') {
            when {
                branch 'main'
            }
            steps {
                echo 'Deploying to PRODUCTION environment'

                // Replace IMAGE_TAG in PROD deployment manifest
                bat """
                powershell -Command "(Get-Content k8s/prod/deployment.yaml) `
                -replace 'IMAGE_TAG', '${IMAGE_TAG}' |
                Set-Content k8s/prod/deployment.yaml"
                """

                // Create namespace if not exists
                bat 'kubectl get namespace prod || kubectl create namespace prod'

                // Apply manifests
                bat 'kubectl apply -f k8s/prod/ --validate=false'
            }
        }
    }

    post {
        success {
            echo "Pipeline SUCCESS for branch: ${BRANCH_NAME}"
            echo "Docker Image Built: ${IMAGE_NAME}"
        }
        failure {
            echo "Pipeline FAILED for branch: ${BRANCH_NAME}"
        }
        always {
            cleanWs()
        }
    }
}



