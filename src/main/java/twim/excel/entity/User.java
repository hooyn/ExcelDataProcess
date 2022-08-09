package twim.excel.entity;

import lombok.Getter;
import lombok.NoArgsConstructor;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;

@Entity
@Getter
@NoArgsConstructor
public class User {
    @Id @GeneratedValue
    private Long id;

    private String username;
    private String age;
    private String tel;

    public User(String username, String age, String tel) {
        this.username = username;
        this.age = age;
        this.tel = tel;
    }
}
