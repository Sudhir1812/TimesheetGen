pipeline {
    agent any

    environment {
        APP_NAME        = "timesheetgen"
        IMAGE_TAG       = "${BRANCH_NAME}-${BUILD_NUMBER}"
        LOCAL_IMAGE     = "${APP_NAME}:${IMAGE_TAG}"
        DOCKERHUB_REPO  = "sudhirkumar181297/timesheetgen"
        REMOTE_IMAGE    = "${DOCKERHUB_REPO}:${IMAGE_TAG}"
        KUBECONFIG      = "C:\\ProgramData\\Jenkins\\.kube\\config"
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
                echo "Building Docker image: ${LOCAL_IMAGE}"
                bat "docker build -t %LOCAL_IMAGE% ."
            }
        }

        stage('Docker Push to Docker Hub') {
            when {
                anyOf {
                    branch 'develop'
                    branch 'main'
                }
            }
            steps {
                withCredentials([usernamePassword(
                    credentialsId: 'dockerhub-creds',
                    usernameVariable: 'DOCKER_USER',
                    passwordVariable: 'DOCKER_PASS'
                )]) {
                    bat """
                        echo %DOCKER_PASS% | docker login -u %DOCKER_USER% --password-stdin
                        docker tag %LOCAL_IMAGE% %REMOTE_IMAGE%
                        docker push %REMOTE_IMAGE%
                        docker logout
                    """
                }
            }
        }

        stage('Deploy to TEST') {
            when {
                branch 'develop'
            }
            steps {
                echo 'Deploying to TEST environment'

                bat """
                powershell -Command "(Get-Content k8s/test/deployment.yaml) -replace 'IMAGE_TAG', '${IMAGE_TAG}' | Set-Content k8s/test/deployment.yaml"
                """

                bat 'kubectl get namespace test || kubectl create namespace test'
                bat 'kubectl apply -f k8s/test/ --validate=false'
            }
        }

        stage('Deploy to PROD') {
            when {
                branch 'main'
            }
            steps {
                echo 'Deploying to PRODUCTION environment'

                bat """
                powershell -Command "(Get-Content k8s/prod/deployment.yaml) -replace 'IMAGE_TAG', '${IMAGE_TAG}' | Set-Content k8s/prod/deployment.yaml"
                """

                bat 'kubectl get namespace prod || kubectl create namespace prod'
                bat 'kubectl apply -f k8s/prod/ --validate=false'
            }
        }
    }

    post {
        success {
            echo "Pipeline SUCCESS for branch: ${BRANCH_NAME}"
            echo "Docker Image Pushed: ${REMOTE_IMAGE}"
        }
        failure {
            echo "Pipeline FAILED for branch: ${BRANCH_NAME}"
        }
        always {
            cleanWs()
        }
    }
}
