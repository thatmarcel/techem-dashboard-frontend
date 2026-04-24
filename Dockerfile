FROM almalinux:10

RUN dnf install -y unzip && curl -fsSL https://bun.sh/install | bash
ENV PATH="$PATH:/root/.bun/bin"

WORKDIR /app

COPY . .

ENTRYPOINT [ "bun", "x", "serve", "-p", "3000" ]
